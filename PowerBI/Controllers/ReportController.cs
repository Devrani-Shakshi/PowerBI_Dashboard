using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Services;
using PowerBI.Data;
using System.Diagnostics;
using Microsoft.PowerBI.Api.Models;

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
                await _service.SyncReports(workspaceId, Guid.Parse(ws.PowerBIWorkspaceId), _db);
            }
            return RedirectToAction("Index", new { workspaceId });
        }

        private async Task<List<PowerBI.Models.Report>> GetReportsList(int workspaceId)
        {
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) 
                return new List<PowerBI.Models.Report>();

            try
            {
                await _service.SyncReports(workspaceId, Guid.Parse(ws.PowerBIWorkspaceId), _db);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[REPORT] Sync Error: {ex.Message}");
                ViewBag.SyncError = "Failed to sync reports: " + ex.Message;
            }

            var reports = _db.Reports
                .Where(r => r.WorkspaceId == workspaceId)
                .OrderBy(r => r.FolderId)
                .ToList();

            var folders = _db.Folders
                .Where(f => f.WorkspaceId == workspaceId)
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
                var import = await _service.UploadReport(
                    workspaceId,
                    Guid.Parse(ws.PowerBIWorkspaceId),
                    file.FileName,
                    stream,
                    _db,
                    folderName,
                    customParams);

                if (import.Reports == null || !import.Reports.Any())
                {
                    return BadRequest("Import failed.");
                }

                var pbiReportId = import.Reports.First().Id;
                
                // Look up folder ID
                int? folderId = null;
                if (!string.IsNullOrEmpty(folderName))
                    folderId = await _service.GetOrCreateFolder(workspaceId, Guid.Parse(ws.PowerBIWorkspaceId), folderName, _db);

                var newReport = new PowerBI.Models.Report
                {
                    Name = file.FileName.Replace(" ", "_").Replace("(", "").Replace(")", ""),
                    PowerBIReportId = pbiReportId.ToString(),
                    WorkspaceId = workspaceId,
                    FilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", file.FileName.Replace(" ", "_").Replace("(", "").Replace(")", "")),
                    FolderId = folderId,
                    ReportType = file.FileName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? "RDL" : "PowerBI"
                };

                _db.Reports.Add(newReport);
                await _db.SaveChangesAsync();

                // Save Custom Filters to DB so they show up in UI immediately
                foreach (var field in customParams)
                {
                    _db.ReportFilters.Add(new ReportFilter
                    {
                        ReportId = newReport.Id,
                        ColumnName = field,
                        DisplayName = field,
                        IsActive = true,
                        IsCustom = true,
                        TableName = "Custom"
                    });
                }
                
                // Add TenantId as a background system filter if it's an RDL
                if (newReport.ReportType == "RDL")
                {
                    _db.ReportFilters.Add(new ReportFilter { ReportId = newReport.Id, ColumnName = "TenantId", DisplayName = "Tenant ID", IsActive = true, IsCustom = false, TableName = "System" });
                }

                await _db.SaveChangesAsync();
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
                var pdfStream = await _service.ExportReportAsStream(
                    Guid.Parse(ws.PowerBIWorkspaceId),
                    Guid.Parse(report.PowerBIReportId),
                    filters,
                    report.ReportType,
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
            Console.WriteLine($"[REPORT] Generating Preview for Report ID: {reportId}");
            
            var report = _db.Reports.Find(reportId);
            if (report == null || string.IsNullOrEmpty(report.PowerBIReportId)) return NotFound();

            var ws = _db.Workspaces.Find(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            var embedConfig = await _service.GetEmbedConfig(
                Guid.Parse(ws.PowerBIWorkspaceId),
                Guid.Parse(report.PowerBIReportId));
            
            embedConfig.ReportType = report.ReportType;
            embedConfig.LocalReportId = report.Id;
            embedConfig.WorkspaceId = ws.PowerBIWorkspaceId;

            Console.WriteLine($"[REPORT] Embed Config generated for: {report.Name} ({report.ReportType})");
            return View(embedConfig);
        }

        [HttpPost]
        public async Task<IActionResult> ResetFilters(int reportId)
        {
            var report = _db.Reports.Find(reportId);
            if (report == null) return NotFound();

            // 1. Purge all existing metadata
            var existing = _db.ReportFilters.Where(f => f.ReportId == reportId).ToList();
            Console.WriteLine($"[CONTROLLER] EXPLICIT RESET: DELETING {existing.Count} ENTRIES FOR REPORT {reportId}");
            
            _db.ReportFilters.RemoveRange(existing);
            await _db.SaveChangesAsync();

            // 2. Force fresh discovery
            try
            {
                Guid? dsId = !string.IsNullOrEmpty(report.PowerBIDatasetId) ? Guid.Parse(report.PowerBIDatasetId) : (Guid?)null;
                await _service.DiscoverReportFilters(reportId, dsId, _db);
                return Ok(new { message = "Schema discovered and saved successfully." });
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
            var report = await _db.Reports.FindAsync(req.ReportId);
            if (report == null) return NotFound();

            var workspace = await _db.Workspaces.FindAsync(report.WorkspaceId);
            if (workspace == null) return BadRequest("Workspace not found");

            Console.WriteLine($"[CONTROLLER] Adding custom filter '{req.DisplayName}' for report '{report.Name}'");

            // 1. If RDL, we must inject into the XML and RE-UPLOAD
            if (report.ReportType == "RDL")
            {
                try 
                {
                    Console.WriteLine($"[STEP 1/4] Starting RDL XML Injection for filter: {req.ColumnName}");
                    await _service.InjectManualFilterToRdl(report.Id, req.ColumnName, _db);
                    Console.WriteLine($"[STEP 2/4] Injection complete. Preparing re-upload for: {report.Name}");

                    // Save old Report ID so we can delete the stale copy later
                    var oldPbiReportId = report.PowerBIReportId;

                    // Re-upload to Power BI to replace existing
                    var fileName = report.Name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? report.Name : $"{report.Name}.rdl";
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);
                    
                    if (string.IsNullOrEmpty(workspace.PowerBIWorkspaceId))
                        throw new Exception("Power BI Workspace ID is missing.");

                    // Small delay to let OS release file handle if necessary
                    await Task.Delay(500);

                    Console.WriteLine($"[STEP 3/4] Opening file stream for re-upload: {filePath}");
                    Import importResult;
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        Console.WriteLine($"[STEP 4/4] Uploading to Power BI workspace: {workspace.PowerBIWorkspaceId}");
                        importResult = await _service.UploadReport(
                            report.WorkspaceId,
                            Guid.Parse(workspace.PowerBIWorkspaceId),
                            fileName,
                            stream,
                            _db,
                            "Automated Reports");
                    }

                    // CRITICAL: Update local DB with the NEW report ID from Power BI
                    if (importResult?.Reports != null && importResult.Reports.Any())
                    {
                        var newReportId = importResult.Reports.First().Id.ToString();
                        Console.WriteLine($"[STEP 5/5] Updating local DB: {oldPbiReportId} -> {newReportId}");
                        report.PowerBIReportId = newReportId;
                        await _db.SaveChangesAsync();
                    }
                    else
                    {
                        Console.WriteLine("[WARNING] Import succeeded but no report ID returned. Embed may use stale ID.");
                    }

                    Console.WriteLine("[SUCCESS] RDL Filter injected and report re-uploaded to Power BI.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERROR] RDL Filter process failed: {ex.Message}");
                    return BadRequest(new { message = $"RDL Update Failed: {ex.Message}" });
                }
            }

            // 2. Save metadata to DB so it shows in sidebar
            var newFilter = new PowerBI.Models.ReportFilter
            {
                ReportId = req.ReportId,
                TableName = req.TableName ?? (report.ReportType == "RDL" ? "RDL_PARAMETER" : "Custom"),
                ColumnName = req.ColumnName,
                DisplayName = req.DisplayName,
                IsActive = true
            };

            _db.ReportFilters.Add(newFilter);
            await _db.SaveChangesAsync();

            return Ok(new { message = "Filter added successfully." });
        }

        public class CustomFilterRequest
        {
            public int ReportId { get; set; }
            public string DisplayName { get; set; }
            public string TableName { get; set; }
            public string ColumnName { get; set; }
        }

        [HttpGet]
        public async Task<IActionResult> GetFilters(int reportId)
        {
            var filters = _db.ReportFilters
                .Where(f => f.ReportId == reportId && f.IsActive)
                .ToList();
            
            return Json(filters);
        }

        [HttpPost]
        public async Task<IActionResult> SaveManualFilter(int reportId, string displayName, string tableName, string columnName)
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
                return BadRequest("Table and Column names are required.");

            var filter = new ReportFilter
            {
                ReportId = reportId,
                DisplayName = displayName ?? columnName,
                TableName = tableName,
                ColumnName = columnName,
                IsActive = true
            };

            _db.ReportFilters.Add(filter);
            await _db.SaveChangesAsync();

            // If it's an RDL report, permanently inject the filter into the XML and re-upload
            var report = await _db.Reports.FindAsync(reportId);
            if (report != null && report.ReportType == "RDL")
            {
                await _service.InjectManualFilterToRdl(reportId, columnName, _db);
            }

            return Ok(new { message = "Filter added successfully." });
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
        public async Task<IActionResult> GetFilterValues(string? datasetId, string table, string column, string? reportId, string? workspaceId)
        {
            if (string.IsNullOrEmpty(table) || string.IsNullOrEmpty(column))
                return BadRequest("Missing parameters");

            Guid? dsGuid = !string.IsNullOrEmpty(datasetId) ? Guid.Parse(datasetId) : (Guid?)null;
            Guid? repGuid = !string.IsNullOrEmpty(reportId) ? Guid.Parse(reportId) : (Guid?)null;
            Guid? wsGuid = !string.IsNullOrEmpty(workspaceId) ? Guid.Parse(workspaceId) : (Guid?)null;

            if (repGuid.HasValue) 
                Console.WriteLine($">>>> [TERMINAL-LOG] Fetching values for Parameter: '{column}' (Report: {reportId})");

            var values = await _service.GetColumnValues(dsGuid, table, column, repGuid, wsGuid, _db);
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

    }
}
