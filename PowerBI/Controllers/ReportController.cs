using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Services;
using PowerBI.Data;
using System.Diagnostics;
using Microsoft.PowerBI.Api.Models;
using Microsoft.EntityFrameworkCore;
using System.Text;

namespace PowerBI.Controllers
{
    public class ReportController : Controller
    {
        private readonly PowerBIService _service;
        private readonly AppDbContext _db;

        public ReportController(PowerBIService service, AppDbContext db)
        {
            _service = service;
            _db = db;
        }

        public async Task<IActionResult> Index(int workspaceId)
        {
            return View(await GetReportsList(workspaceId));
        }

        public async Task<IActionResult> Sync(int workspaceId)
        {
            Console.WriteLine($"[REPORT] FORCED SYNC for Workspace ID: {workspaceId}");
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId))
            {
                await _service.SyncReports(new List<int> { workspaceId }, _db);
            }
            return RedirectToAction("Index", new { workspaceId });
        }

        private async Task<List<PowerBI.Models.Report>> GetReportsList(int workspaceId)
        {
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) 
                return new List<PowerBI.Models.Report>();

            var allRelatedWorkspaceIds = _db.Workspaces
                .Where(w => w.PowerBIWorkspaceId == ws.PowerBIWorkspaceId)
                .Select(w => w.Id)
                .ToList();

            try
            {
                await _service.SyncReports(allRelatedWorkspaceIds, _db);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[REPORT] Sync Error: {ex.Message}");
                ViewBag.SyncError = "Failed to sync reports: " + ex.Message;
            }

            var reports = _db.Reports
                .Where(r => allRelatedWorkspaceIds.Contains(r.WorkspaceId))
                .AsEnumerable()
                .GroupBy(r => $"{r.Name}|{r.FolderId}") // Deduplicate by Name + Folder
                .Select(g => g.OrderByDescending(r => r.Id).First()) // Pick the latest local record
                .OrderBy(r => r.FolderId)
                .ToList();

            var folders = _db.Folders
                .Where(f => allRelatedWorkspaceIds.Contains(f.WorkspaceId))
                .AsEnumerable()
                .GroupBy(f => f.FabricFolderId)
                .Select(g => g.First()) // Deduplicate
                .ToList();

            ViewBag.WorkspaceId = workspaceId;
            ViewBag.WorkspaceName = ws.Name;
            ViewBag.Folders = folders;
            return reports;
        }

        public async Task<IActionResult> Upload(IFormFile file, int workspaceId, string? folderName, string? targetFields)
        {
            Console.WriteLine($"[REPORT] Uploading file: {file.FileName} to Workspace: {workspaceId} (Fields: {targetFields})");
            
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) 
            {
                return BadRequest("Invalid Workspace");
            }

            // Parse target fields
            var customParams = targetFields?.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList() ?? new List<string>();

            try 
            {
                using var stream = file.OpenReadStream();
                var uploadResult = await _service.UploadReport(
                    workspaceId,
                    Guid.Parse(ws.PowerBIWorkspaceId),
                    file.FileName,
                    stream,
                    _db,
                    folderName,
                    customParams);

                var import = uploadResult.Import;
                var rdlContent = uploadResult.RdlContent;

                if (import.Reports == null || !import.Reports.Any())
                {
                    return BadRequest("Import failed.");
                }

                var pbiReportId = import.Reports.First().Id;

                int? folderId = null;
                if (!string.IsNullOrEmpty(folderName))
                    folderId = await _service.GetOrCreateFolder(workspaceId, Guid.Parse(ws.PowerBIWorkspaceId), folderName, _db);

                var currentUserId = HttpContext.Session.GetInt32("UserId");

                // DEDUPLICATION: Check if a report with this name already exists in this workspace/folder
                var normalizedName = file.FileName.Replace(" ", "_").Replace("(", "").Replace(")", "");
                var existingReport = _db.Reports.FirstOrDefault(r => 
                    r.WorkspaceId == workspaceId && 
                    r.FolderId == folderId && 
                    r.Name == normalizedName);

                PowerBI.Models.Report reportRecord;
                if (existingReport != null)
                {
                    Console.WriteLine($"[UPLOAD] Updating existing report record: {normalizedName} (ID: {existingReport.Id})");
                    existingReport.PowerBIReportId = pbiReportId.ToString();
                    existingReport.RdlContent = rdlContent;
                    reportRecord = existingReport;
                }
                else
                {
                    Console.WriteLine($"[UPLOAD] Creating new report record: {normalizedName}");
                    reportRecord = new PowerBI.Models.Report
                    {
                        Name = normalizedName,
                        PowerBIReportId = pbiReportId.ToString(),
                        WorkspaceId = workspaceId,
                        FilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", normalizedName),
                        FolderId = folderId,
                        CreatedByUserId = currentUserId,
                        ReportType = file.FileName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? "RDL" : "PowerBI",
                        RdlContent = rdlContent
                    };
                    _db.Reports.Add(reportRecord);
                }
                
                await _db.SaveChangesAsync();

                // Save Custom Filters to DB so they show up in UI immediately
                foreach (var field in customParams)
                {
                    // Check for duplicate filters before adding
                    var existingFilter = _db.ReportFilters.FirstOrDefault(f => f.ReportId == reportRecord.Id && f.ColumnName == field && f.UserId == currentUserId);
                    if (existingFilter == null)
                    {
                        _db.ReportFilters.Add(new ReportFilter
                        {
                            ReportId = reportRecord.Id,
                            ColumnName = field,
                            DisplayName = field,
                            IsActive = true,
                            IsCustom = true,
                            TableName = "Custom",
                            UserId = currentUserId
                        });
                    }
                }
                
                // Add TenantId as a background system filter if it's an RDL
                if (reportRecord.ReportType == "RDL")
                {
                    _db.ReportFilters.Add(new ReportFilter { ReportId = reportRecord.Id, ColumnName = "TenantId", DisplayName = "Tenant ID", IsActive = true, IsCustom = false, TableName = "System", UserId = currentUserId });
                }

                await _db.SaveChangesAsync();

                // --- CRITICAL SYNC: Re-inject all existing filters into the newly uploaded RDL ---
                if (reportRecord.ReportType == "RDL")
                {
                    var existingFilters = _db.ReportFilters.Where(f => f.ReportId == reportRecord.Id && f.IsCustom).ToList();
                    Console.WriteLine($"[UPLOAD-SYNC] Detected {existingFilters.Count} existing custom filters. Triggering automatic RDL injection...");
                    
                    foreach (var filter in existingFilters)
                    {
                        var (injectSuccess, injectMsg) = await _service.InjectFilterToRdl(reportRecord.Id, filter.ColumnName, _db);
                        Console.WriteLine($"[UPLOAD-SYNC] Injecting '{filter.ColumnName}': {injectMsg}");
                    }
                }

                return RedirectToAction("Index", new { workspaceId });
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Contains("Unauthorized"))
                {
                    message = "Unauthorized. Please ensure the Service Principal (App Registration) is added as an ADMIN or MEMBER to this specific workspace in the Power BI portal.";
                }
                else if (message.Contains("Forbidden"))
                {
                    message = "Forbidden. This usually means the Service Principal is not an admin of the workspace, or 'Export to PDF' is disabled for SPs in Tenant settings.";
                }
                else if (message.Contains("RequestedFileIsEncryptedOrCorrupted"))
                {
                    bool isRdl = file.FileName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase);
                    message = "The file appears to be corrupted, encrypted, or not a valid Power BI report.";
                    if (isRdl)
                    {
                        message += " IMPORTANT: Paginated Reports (.rdl) require the target workspace to be on a Premium or Fabric capacity.";
                    }
                    message += " Also ensure the file is not protected by sensitivity labels.";
                }
                TempData["Error"] = "Upload Error: " + message;
                return RedirectToAction("Index", new { workspaceId });
            }
        }

        public async Task<IActionResult> Export(int reportId)
        {
            return await ExportWithFilters(reportId, null);
        }

        [HttpPost]
        public async Task<IActionResult> ExportWithFilters(int reportId, [FromBody] List<ExportFilter>? filters)
        {
            Console.WriteLine($"[REPORT] Exporting PDF for Report ID: {reportId} (Filters: {filters?.Count ?? 0})");
            
            var report = _db.Reports.Find(reportId);
            if (report == null || string.IsNullOrEmpty(report.PowerBIReportId)) return NotFound();

            var ws = _db.Workspaces.Find(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            try
            {
                string reportType = report.ReportType ?? "PowerBI";
                var pdfStream = await _service.ExportReportAsStream(
                    Guid.Parse(ws.PowerBIWorkspaceId),
                    Guid.Parse(report.PowerBIReportId),
                    filters,
                    reportType,
                    _db);

                var fileName = $"{report.Name}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
                var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadsDir)) Directory.CreateDirectory(uploadsDir);
                
                var filePath = Path.Combine(uploadsDir, fileName);

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    await pdfStream.CopyToAsync(fileStream);
                }

                Console.WriteLine($"[REPORT] PDF saved to local server: {filePath}");
                
                var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[REPORT] Export Error: {ex.Message}");
                return BadRequest(new { message = "Export failed. Ensure the workspace is on a Fabric/Premium capacity and the Service Principal has 'Export to File' permissions." });
            }
        }

        public async Task<IActionResult> Preview(int reportId)
        {
            var currentUserId = HttpContext.Session.GetInt32("UserId");
            if (currentUserId == null) return RedirectToAction("Login", "Auth");

            Console.WriteLine($"[SINGLE-REPORT] Generating Preview for Report ID: {reportId} (User: {currentUserId})");
            
            var report = _db.Reports.Find(reportId);
            if (report == null || string.IsNullOrEmpty(report.PowerBIReportId)) return NotFound();

            var ws = _db.Workspaces.Find(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            // ALWAYS USE THE ORIGINAL REPORT ID
            string finalPbiReportId = report.PowerBIReportId;

            var embedConfig = await _service.GetEmbedConfig(
                Guid.Parse(ws.PowerBIWorkspaceId),
                Guid.Parse(finalPbiReportId),
                _db);
            
            embedConfig.ReportType = report.ReportType;
            embedConfig.LocalReportId = report.Id;
            embedConfig.WorkspaceId = ws.PowerBIWorkspaceId;

            return View(embedConfig);
        }
        public async Task<IActionResult> ResetFilters(int reportId)
        {
            var report = _db.Reports.Find(reportId);
            if (report == null) return NotFound();

            var currentUserId = HttpContext.Session.GetInt32("UserId");
            if (currentUserId == null) return Unauthorized("Session expired. Please log in again.");
            
            Console.WriteLine($"[SINGLE-REPORT] SYNCING RDL PARAMETERS FOR REPORT {reportId} (USER {currentUserId})");

            try
            {
                // Step 1: [AZURE-IT-SPEC] - Patch Credentials & Gateway Binding FIRST
                if (report.ReportType == "RDL" && !string.IsNullOrEmpty(report.PowerBIReportId))
                {
                    Console.WriteLine(">>>> [FLOW] RDL detected. Establishing Gateway Connectivity FIRST...");
                    var ws = _db.Workspaces.Find(report.WorkspaceId);
                    if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId))
                    {
                        await _service.PatchReportCredentials(Guid.Parse(ws.PowerBIWorkspaceId), Guid.Parse(report.PowerBIReportId));
                    }
                }

                // Step 2: Metadata Discovery (Returns true if self-healing injection occurred)
                Console.WriteLine($">>>> [FLOW] Starting Metadata Discovery Process for Report {reportId}...");
                Guid? dsId = !string.IsNullOrEmpty(report.PowerBIDatasetId) ? Guid.Parse(report.PowerBIDatasetId) : (Guid?)null;
                bool wasModified = await _service.DiscoverReportFilters(reportId, dsId, _db, currentUserId.Value);

                // Step 3: Cloud Sync (If discovery modified the RDL XML)
                if (wasModified && report.ReportType == "RDL")
                {
                    Console.WriteLine(">>>> [FLOW] RDL was modified during discovery. Syncing changes to cloud...");
                    var ws = await _db.Workspaces.FindAsync(report.WorkspaceId);
                    if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId))
                    {
                        string fileName = report.Name.EndsWith(".rdl") ? report.Name : $"{report.Name}.rdl";
                        byte[] fileBytes = Encoding.UTF8.GetBytes(report.RdlContent ?? "");
                        using var ms = new MemoryStream(fileBytes);
                        
                        var (uploadResult, _) = await _service.UploadReport(report.WorkspaceId, Guid.Parse(ws.PowerBIWorkspaceId), fileName, ms, _db);
                        
                        // Update to the NEW ID generated by replacement
                        if (uploadResult.Reports != null && uploadResult.Reports.Any())
                        {
                            report.PowerBIReportId = uploadResult.Reports.First().Id.ToString();
                            report.PowerBIDatasetId = uploadResult.Reports.First().DatasetId;
                            await _db.SaveChangesAsync();
                        }
                    }
                }

                Console.WriteLine(">>>> [FLOW] Filter Discovery and Sync COMPLETED.");
                var filters = await _db.ReportFilters
                    .Where(f => f.ReportId == reportId && f.IsActive && (f.UserId == null || f.UserId == currentUserId))
                    .OrderBy(f => f.IsCustom)
                    .ToListAsync();

                return Ok(new 
                { 
                    message = "Parameters synchronized successfully.", 
                    filters = filters,
                    reportId = report.PowerBIReportId
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CONTROLLER] Discovery failed: {ex.Message}");
                return BadRequest(new { message = $"Discovery failed: {ex.Message}" });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SaveCustomFilter([FromBody] CustomFilterRequest req)
        {
            try
            {
                if (req == null) return BadRequest(new { success = false, message = "Invalid request body." });

                var currentUserId = HttpContext.Session.GetInt32("UserId");
                if (currentUserId == null) return Unauthorized(new { success = false, message = "Session expired. Please login again." });

                var report = await _db.Reports.FindAsync(req.ReportId);
                if (report == null) return NotFound(new { success = false, message = "Report not found." });

                if (string.IsNullOrEmpty(req.ColumnName))
                    return BadRequest(new { success = false, message = "Field/Column Name is required." });

                // --- SMART SETUP: One-time RDL Injection ---
                if (report.ReportType == "RDL")
                {
                    var validation = await _service.ValidateFieldExistsInRdl(report.Id, req.ColumnName, _db);
                    if (validation.Exists)
                    {
                        Console.WriteLine($"[SMART-SETUP] Injecting parameter for '{req.ColumnName}'...");
                        var injectionResult = await _service.InjectFilterToRdl(report.Id, req.ColumnName, _db);
                        if (injectionResult.Success && injectionResult.Message == "Success")
                        {
                            // Push the updated RDL to Power BI (One-time sync)
                            var workspace = await _db.Workspaces.FindAsync(report.WorkspaceId);
                            if (workspace != null && !string.IsNullOrEmpty(workspace.PowerBIWorkspaceId))
                            {
                                string reportName = report.Name ?? "UnknownReport";
                                string fileName = reportName.EndsWith(".rdl") ? reportName : $"{reportName}.rdl";
                                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);
                                byte[] fileBytes;
                                if (!string.IsNullOrEmpty(report.RdlContent))
                                {
                                    fileBytes = Encoding.UTF8.GetBytes(report.RdlContent);
                                    Console.WriteLine("[SMART-SETUP] Using Injected XML from Database for Cloud Sync.");
                                }
                                else
                                {
                                    fileBytes = System.IO.File.ReadAllBytes(filePath);
                                    Console.WriteLine("[SMART-SETUP] Warning: RdlContent was null. Falling back to physical file.");
                                }

                                 using var memoryStream = new MemoryStream(fileBytes);
                                var (uploadResult, injectedXml) = await _service.UploadReport(report.WorkspaceId, Guid.Parse(workspace.PowerBIWorkspaceId), fileName, memoryStream, _db);
                                
                                // CRITICAL: Update the report ID and Dataset ID in local DB because the old one was deleted for replacement
                                if (uploadResult != null && uploadResult.Reports != null && uploadResult.Reports.Any())
                                {
                                    var newReport = uploadResult.Reports.First();
                                    report.PowerBIReportId = newReport.Id.ToString();
                                    report.PowerBIDatasetId = newReport.DatasetId;
                                    Console.WriteLine($"[SMART-SETUP] Database updated with NEW Report ID: {report.PowerBIReportId}");
                                }
                                
                                Console.WriteLine("[SMART-SETUP] RDL updated in cloud. Parameter is now ready for Native Apply.");
                            }
                        }
                    }
                }

                string displayName = req.DisplayName ?? req.ColumnName;
                string tableName = req.TableName ?? (report.ReportType == "RDL" ? "RDL_PARAMETER" : "Custom");

                var existing = await _db.ReportFilters.FirstOrDefaultAsync(f => 
                    f.ReportId == req.ReportId && 
                    f.TableName == tableName && 
                    f.ColumnName == req.ColumnName && 
                    f.UserId == currentUserId);

                if (existing != null)
                {
                    existing.IsActive = true;
                    existing.DisplayName = displayName;
                    existing.IsCustom = true;
                }
                else
                {
                    var newFilter = new PowerBI.Models.ReportFilter
                    {
                        ReportId = req.ReportId,
                        TableName = tableName,
                        ColumnName = req.ColumnName,
                        DisplayName = displayName,
                        IsActive = true,
                        IsCustom = true,
                        UserId = currentUserId
                    };
                    _db.ReportFilters.Add(newFilter);
                }

                await _db.SaveChangesAsync();
                return Ok(new { success = true, message = "Filter saved successfully." });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] SaveCustomFilter Failed: {ex.Message}");
                return StatusCode(500, new { success = false, message = "Error saving filter.", detail = ex.Message });
            }
        }

        public class CustomFilterRequest
        {
            public int ReportId { get; set; }
            public string? DisplayName { get; set; }
            public string? TableName { get; set; }
            public string? ColumnName { get; set; }
        }

        [HttpGet]
        public async Task<IActionResult> GetFilters(int reportId)
        {
            var currentUserId = HttpContext.Session.GetInt32("UserId");
            Console.WriteLine($"[GET-FILTERS] Request for Report {reportId} from User {currentUserId}");

            var filters = _db.ReportFilters
                .Where(f => f.ReportId == reportId && f.IsActive && (f.UserId == null || f.UserId == currentUserId))
                .OrderBy(f => f.IsCustom)
                .ToList();
            
            Console.WriteLine($"[GET-FILTERS] Returning {filters.Count} filters for user {currentUserId}.");
            foreach(var f in filters) {
                Console.WriteLine($" -> Filter: {f.DisplayName} (Table: {f.TableName}, Column: {f.ColumnName}, Owner: {f.UserId})");
            }
            
            return Json(filters);
        }

        public class BulkFilterRequest
        {
            public int ReportId { get; set; }
            public List<CustomFilterRequest> Filters { get; set; } = new();
        }

        [HttpPost]
        public async Task<IActionResult> SaveBulkFilters([FromBody] BulkFilterRequest req)
        {
            var report = await _db.Reports.FindAsync(req.ReportId);
            if (report == null) return NotFound();

            var currentUserId = HttpContext.Session.GetInt32("UserId");
            int addedCount = 0;

            foreach (var fReq in req.Filters)
            {
                if (string.IsNullOrEmpty(fReq.ColumnName)) continue;

                var existing = await _db.ReportFilters.FirstOrDefaultAsync(f => 
                    f.ReportId == req.ReportId && 
                    f.TableName == fReq.TableName && 
                    f.ColumnName == fReq.ColumnName && 
                    f.UserId == currentUserId);

                if (existing == null)
                {
                    _db.ReportFilters.Add(new ReportFilter
                    {
                        ReportId = req.ReportId,
                        TableName = fReq.TableName ?? "Custom",
                        ColumnName = fReq.ColumnName,
                        DisplayName = fReq.DisplayName ?? fReq.ColumnName,
                        IsActive = true,
                        IsCustom = true,
                        UserId = currentUserId
                    });
                    addedCount++;
                }
                else if (!existing.IsActive)
                {
                    existing.IsActive = true;
                    addedCount++;
                }
            }

            await _db.SaveChangesAsync();
            return Ok(new { message = $"Saved {addedCount} filters." });
        }

        [HttpPost]
        public async Task<IActionResult> Delete(int reportId)
        {
            var report = await _db.Reports.FindAsync(reportId);
            if (report == null) return NotFound();

            var workspaceId = report.WorkspaceId;
            var ws = await _db.Workspaces.FindAsync(workspaceId);

            try 
            {
                var currentUserId = HttpContext.Session.GetInt32("UserId");
                var userRole = HttpContext.Session.GetString("UserRole");
                
                Console.WriteLine($"[REPORT-DELETE] Attempt by User {currentUserId} (Role: {userRole}) for Report {reportId}");

                // Authorization: Only the owner OR Admin can delete a report
                bool isOwner = report.CreatedByUserId == currentUserId;
                bool isAdmin = userRole == "Admin";
                bool isLegacy = report.CreatedByUserId == null; // Handle old files

                if (!isOwner && !isAdmin && !isLegacy)
                {
                    Console.WriteLine("[REPORT-DELETE] REJECTED: User is not owner and not admin.");
                    TempData["Error"] = "Access Denied: Only the user who uploaded this report or an Admin can delete it.";
                    return RedirectToAction("Index", new { workspaceId });
                }

                if (isAdmin) Console.WriteLine("[REPORT-DELETE] Administrative override granted.");

                if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId) && !string.IsNullOrEmpty(report.PowerBIReportId))
                {
                    await _service.DeleteReport(Guid.Parse(ws.PowerBIWorkspaceId), Guid.Parse(report.PowerBIReportId));
                }
                
                _db.Reports.Remove(report);
                await _db.SaveChangesAsync();
                
                return RedirectToAction("Index", new { workspaceId });
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Delete Failed: " + ex.Message;
                return RedirectToAction("Index", new { workspaceId });
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetFilterValues(string table, string column, int localReportId, string? datasetId)
        {
            if (string.IsNullOrEmpty(table) || string.IsNullOrEmpty(column))
                return BadRequest("Missing parameters");

            var currentUserId = HttpContext.Session.GetInt32("UserId");
            if (currentUserId == null) return Unauthorized();

            var report = await _db.Reports.FindAsync(localReportId);
            if (report == null) return NotFound("Local report not found.");

            var ws = await _db.Workspaces.FindAsync(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            // Always use the base report ID in Single Report Architecture
            string finalPbiReportId = report.PowerBIReportId ?? "";
            if (string.IsNullOrEmpty(finalPbiReportId)) return BadRequest("Report ID is missing.");

            Guid.TryParse(datasetId, out Guid dsGuidParsed);
            Guid? dsGuid = dsGuidParsed == Guid.Empty ? (Guid?)null : dsGuidParsed;
            Guid.TryParse(finalPbiReportId, out Guid repGuid);
            Guid.TryParse(ws.PowerBIWorkspaceId, out Guid wsGuid);
            if (report.ReportType == "RDL" || report.ReportType == "PaginatedReport") dsGuid = null;

            var values = await _service.GetColumnValues(dsGuid, table, column, repGuid, wsGuid, _db, report.ReportType ?? "PowerBI");
            return Json(values);
        }

        [HttpGet]
        public async Task<IActionResult> GetDatasetSchema(string datasetId)
        {
            Console.WriteLine($"[CONTROLLER] GetDatasetSchema called for Dataset ID: {datasetId}");
            if (string.IsNullOrEmpty(datasetId)) 
            {
                Console.WriteLine("[CONTROLLER] ERROR: Dataset ID is null or empty.");
                return BadRequest("Dataset ID is required.");
            }
            
            try 
            {
                var schema = await _service.GetDatasetTablesAndColumns(Guid.Parse(datasetId));
                Console.WriteLine($"[CONTROLLER] Schema discovery successful. Found {schema.Count} fields.");
                return Json(schema);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CONTROLLER] ERROR in GetDatasetSchema: {ex.Message}");
                return BadRequest(new { message = ex.Message });
            }
        }

        [HttpPost]
        public async Task<IActionResult> DeleteFilter(int filterId)
        {
            Console.WriteLine($"[FILTER-DELETE] Request for ID: {filterId}");
            var filter = await _db.ReportFilters.FindAsync(filterId);
            if (filter == null) return NotFound();

            var currentUserId = HttpContext.Session.GetInt32("UserId");
            var userRole = HttpContext.Session.GetString("UserRole");

            if (filter.UserId != currentUserId && userRole != "Admin")
            {
                return BadRequest("Access Denied.");
            }

            filter.IsActive = false;
            await _db.SaveChangesAsync();
            return Ok();
        }

        [HttpPost]
        public async Task<IActionResult> DeleteFolder(int folderId)
        {
            Console.WriteLine($"[FOLDER-DELETE] Request for ID: {folderId}");
            var folder = await _db.Folders.FindAsync(folderId);
            if (folder == null) return NotFound();

            var workspaceId = folder.WorkspaceId;
            var ws = await _db.Workspaces.FindAsync(workspaceId);

            try 
            {
                // 1. Delete from Fabric Service
                if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId) && !string.IsNullOrEmpty(folder.FabricFolderId))
                {
                    await _service.DeleteFabricItem(Guid.Parse(ws.PowerBIWorkspaceId), Guid.Parse(folder.FabricFolderId));
                }

                // 2. Clear local report associations
                var reports = _db.Reports.Where(r => r.FolderId == folderId).ToList();
                foreach(var r in reports) r.FolderId = null;

                // 3. Remove from local DB
                _db.Folders.Remove(folder);
                await _db.SaveChangesAsync();
                
                TempData["Success"] = "Folder deleted successfully from Fabric and local DB.";
            }
            catch (Exception ex) when (ex.Message.Contains("404") || ex.Message.Contains("NotFound"))
            {
                Console.WriteLine($"[FOLDER-DELETE] Fabric item already gone (404). Cleaning up local DB...");
                // Proceed with local cleanup
                var reports = _db.Reports.Where(r => r.FolderId == folderId).ToList();
                foreach(var r in reports) r.FolderId = null;

                _db.Folders.Remove(folder);
                await _db.SaveChangesAsync();
                TempData["Success"] = "Folder was already deleted from Fabric. Local record cleaned up.";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FOLDER-DELETE] ERROR: {ex.Message}");
                TempData["Error"] = "Failed to delete folder from Fabric: " + ex.Message;
            }

            return RedirectToAction("Index", new { workspaceId });
        }
    }
}
