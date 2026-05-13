using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using Microsoft.Rest;
using PowerBI.Models;
using PowerBI.Data;
using System.Net.Http.Headers;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Text;
using Microsoft.EntityFrameworkCore;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace PowerBI.Services
{
    public class PowerBIService
    {
        private readonly PowerBIAuthService _auth;
        private readonly IConfiguration _config;

        public PowerBIService(PowerBIAuthService auth, IConfiguration config)
        {
            _auth = auth;
            _config = config;
        }

        private async Task<PowerBIClient> GetClient()
        {
            Console.WriteLine("[SERVICE] Fetching token from Azure AD...");
            var token = await _auth.GetAccessToken();
            
            // Log first 10 chars of token for debugging (Safe)
            Console.WriteLine($"[SERVICE] Token acquired (Starts with: {token.Substring(0, 10)}...)");

            var credentials = new TokenCredentials(token, "Bearer");
            return new PowerBIClient(new Uri("https://api.powerbi.com/"), credentials);
        }

        // --- SINGLE REPORT ARCHITECTURE: RDL SCHEMA DISCOVERY ---
        public async Task<(bool Exists, string Detail)> ValidateFieldExistsInRdl(int reportId, string fieldName, AppDbContext db)
        {
            var report = await db.Reports.FindAsync(reportId);
            if (report == null) return (false, "Report not found in database.");

            var reportName = report.Name ?? "Unknown";
            var fileName = reportName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? reportName : $"{reportName}.rdl";
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);

            if (!File.Exists(filePath)) return (false, $"Original RDL file '{fileName}' missing from server uploads.");

            try
            {
                XDocument xDoc = XDocument.Load(filePath);
                var ns = xDoc.Root?.GetDefaultNamespace() ?? XNamespace.None;

                // 1. Check Report Parameters
                var hasParam = xDoc.Descendants(ns + "ReportParameter")
                                   .Any(p => p.Attribute("Name")?.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase) == true);
                if (hasParam) return (true, "Found in Report Parameters");

                // 2. Check Dataset Fields
                var hasField = xDoc.Descendants(ns + "Field")
                                   .Any(f => f.Attribute("Name")?.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase) == true);
                if (hasField) return (true, "Found in Dataset Fields");

                // 3. Check Dataset Query Parameters
                var hasQueryParam = xDoc.Descendants(ns + "QueryParameter")
                                        .Any(q => q.Attribute("Name")?.Value.Replace("@", "").Equals(fieldName, StringComparison.OrdinalIgnoreCase) == true);
                if (hasQueryParam) return (true, "Found in Dataset Query Parameters");

                // 3. CRITICAL: Wipe the Layout Grid (Industry Level Workaround)
                // Deleting the layout section forces Power BI to re-generate a perfect layout for all parameters.
                // This prevents 'rsInvalidParameterLayoutCellDefNotEqualsParameterCount' and 'rsInvalidParameterLayoutParametersMissingFromPanel' errors.
                var layout = xDoc.Descendants(ns + "ReportParametersLayout").FirstOrDefault();
                if (layout != null)
                {
                    layout.Remove();
                    Console.WriteLine(">>>> [LAYOUT] Removed fixed grid layout to force dynamic re-generation.");
                }

                // 4. SAVE WITH BOM (Power BI Requirement)
                var utf8WithBom = new System.Text.UTF8Encoding(true);
                using (var sw = new StreamWriter(filePath, false, utf8WithBom))
                {
                    xDoc.Save(sw);
                }

                return (false, $"Field '{fieldName}' not found in RDL schema. Note: For RDL reports, you can only filter by fields that are defined as Report Parameters.");
            }
            catch (Exception ex)
            {
                return (false, $"Error parsing RDL: {ex.Message}");
            }
        }
        public async Task<Microsoft.PowerBI.Api.Models.Group> GetOrCreateWorkspace(string name)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Checking if workspace '{name}' exists...");
                var client = await GetClient();
                var groups = await client.Groups.GetGroupsAsync();
                
                var existingGroup = groups?.Value?.FirstOrDefault(g => g.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                
                if (existingGroup != null)
                {
                    Console.WriteLine($"[SERVICE] Found existing workspace: {existingGroup.Id}");
                    return existingGroup;
                }

                Console.WriteLine("[SERVICE] Not found. Creating new workspace...");
                var newGroup = await client.Groups.CreateGroupAsync(new GroupCreationRequest { Name = name });
                Console.WriteLine($"[SERVICE] New workspace created: {newGroup.Id}");

                // AUTO-ASSIGN TO CAPACITY
                var capacityId = _auth.GetCapacityId(); 
                if (!string.IsNullOrEmpty(capacityId))
                {
                    Console.WriteLine($"[SERVICE] Auto-assigning to capacity: {capacityId}");
                    try {
                        await client.Groups.AssignToCapacityAsync(newGroup.Id, new AssignToCapacityRequest { CapacityId = Guid.Parse(capacityId) });
                        Console.WriteLine("[SERVICE] Successfully assigned to capacity.");
                    } catch (Exception ex) {
                        Console.WriteLine($"[SERVICE] Capacity assignment failed: {ex.Message}. You may need to assign it manually.");
                    }
                }


                var adminEmail = _auth.GetAdminEmail();
                if (!string.IsNullOrEmpty(adminEmail))
                {
                    try 
                    {
                        Console.WriteLine($"[SERVICE] Adding {adminEmail} as Admin to new workspace...");
                        await client.Groups.AddGroupUserAsync(newGroup.Id, new GroupUser 
                        { 
                            Identifier = adminEmail, 
                            GroupUserAccessRight = "Admin", 
                            PrincipalType = "User" 
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SERVICE] Warning: Could not add {adminEmail} as admin: {ex.Message}");
                    }
                }

                return newGroup;
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] ERROR in GetOrCreateWorkspace: {ex.Response.Content}");
                throw;
            }
        }

     
        public async Task SyncWorkspaces(int userId, AppDbContext db)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Syncing workspaces for User ID: {userId}");
                var client = await GetClient();
                var pbiGroups = await client.Groups.GetGroupsAsync();
                var pbiWorkspaces = pbiGroups.Value;

                var localWorkspaces = db.Workspaces.ToList();
                
                var adminEmail = _auth.GetAdminEmail();
                foreach (var pbiWs in pbiWorkspaces)
                {
                    if (!string.IsNullOrEmpty(adminEmail))
                    {
                        try 
                        {
                            await client.Groups.AddGroupUserAsync(pbiWs.Id, new GroupUser 
                            { 
                                Identifier = adminEmail, 
                                GroupUserAccessRight = "Admin", 
                                PrincipalType = "User" 
                            });
                        }
                        catch { }
                    }
                    
                    var existing = localWorkspaces.FirstOrDefault(w => w.PowerBIWorkspaceId == pbiWs.Id.ToString());
                    if (existing == null)
                    {
                        Console.WriteLine($"[SERVICE] Adding new PBI workspace to DB: {pbiWs.Name}");
                        db.Workspaces.Add(new Workspace
                        {
                            Name = pbiWs.Name,
                            PowerBIWorkspaceId = pbiWs.Id.ToString(),
                            UserId = userId
                        });
                    }
                    else if (existing.Name != pbiWs.Name)
                    {
                        Console.WriteLine($"[SERVICE] Updating workspace name in DB: {existing.Name} -> {pbiWs.Name}");
                        existing.Name = pbiWs.Name;
                    }
                }

                foreach (var localWs in localWorkspaces)
                {
                    if (!pbiWorkspaces.Any(pbi => pbi.Id.ToString() == localWs.PowerBIWorkspaceId))
                    {
                        Console.WriteLine($"[SERVICE] Removing workspace from DB (deleted in PBI): {localWs.Name}");
                        db.Workspaces.Remove(localWs);
                    }
                }

                await db.SaveChangesAsync();
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] Power BI API Error during Sync: {ex.Response.Content}");
                throw new Exception($"Power BI API Error: {ex.Response.ReasonPhrase}. Details: {ex.Response.Content}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SERVICE] General Error during Sync: {ex.Message}");
                throw;
            }
        }


        public async Task RenameWorkspace(Guid workspaceId, string newName)
        {
            Console.WriteLine($"[SERVICE] Renaming workspace {workspaceId} to '{newName}'");
            var client = await GetClient();
            await client.Groups.UpdateGroupAsync(workspaceId, new UpdateGroupRequest { Name = newName });
        }


        public async Task DeleteWorkspace(Guid workspaceId)
        {
            Console.WriteLine($"[SERVICE] Deleting workspace {workspaceId}");
            var client = await GetClient();
            await client.Groups.DeleteGroupAsync(workspaceId);
        }




        public async Task SyncFolders(int localWorkspaceId, Guid pbiWorkspaceId, AppDbContext db)
        {
            try
            {
                Console.WriteLine($"[SERVICE] Syncing folders for Workspace: {pbiWorkspaceId}");
                var token = await _auth.GetFabricToken();
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var url = $"https://api.fabric.microsoft.com/v1/workspaces/{pbiWorkspaceId}/folders";
                var response = await client.GetAsync(url);
                if (!response.IsSuccessStatusCode) return;

                var rawJson = await response.Content.ReadAsStringAsync();
                var root = JsonConvert.DeserializeObject<dynamic>(rawJson);
                var folders = root?.value;

                var localFolders = db.Folders.Where(f => f.WorkspaceId == localWorkspaceId).ToList();

                if (folders != null)
                {
                    foreach (var f in folders)
                    {
                        string? fId = f.id?.ToString();
                        string? fName = f.displayName?.ToString();
                        if (string.IsNullOrEmpty(fId) || string.IsNullOrEmpty(fName)) continue;
                        
                        Console.WriteLine($"[SERVICE] Folder Sync: Found '{fName}' (Fabric ID: {fId})");

                        var existing = localFolders.FirstOrDefault(lf => lf.FabricFolderId == fId);
                        if (existing == null)
                        {
                            db.Folders.Add(new Folder
                            {
                                Name = fName,
                                FabricFolderId = fId,
                                WorkspaceId = localWorkspaceId
                            });
                        }
                        else if (existing.Name != fName)
                        {
                            existing.Name = fName;
                        }
                    }
                }

                // Cleanup deleted folders
                foreach (var lf in localFolders)
                {
                    bool stillExists = false;
                    if (folders != null)
                    {
                        foreach (var f in folders)
                        {
                            if (f.id?.ToString() == lf.FabricFolderId) { stillExists = true; break; }
                        }
                    }
                    if (!stillExists) db.Folders.Remove(lf);
                }

                await db.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SERVICE] SyncFolders Error: {ex.Message}");
            }
        }

        public async Task SyncReports(IEnumerable<int> localWorkspaceIds, AppDbContext db)
        {
            var processedPbiIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var localWorkspaceId in localWorkspaceIds)
            {
                var workspace = await db.Workspaces.FindAsync(localWorkspaceId);
                if (workspace == null || string.IsNullOrEmpty(workspace.PowerBIWorkspaceId)) continue;
                Guid pbiWorkspaceId;
                if (!Guid.TryParse(workspace.PowerBIWorkspaceId, out pbiWorkspaceId)) continue;

                // 1. Sync Folders First
                await SyncFolders(localWorkspaceId, pbiWorkspaceId, db);

                Console.WriteLine($"[SERVICE] Syncing reports for Workspace: {pbiWorkspaceId}");
                var client = await GetClient();
                var pbiReports = (await client.Reports.GetReportsInGroupAsync(pbiWorkspaceId)).Value;

                // 2. Fetch ALL items from Fabric (MUST be recursive to see parentFolderId correctly)
                var fabricToken = await _auth.GetFabricToken();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", fabricToken);
                
                var itemsUrl = $"https://api.fabric.microsoft.com/v1/workspaces/{pbiWorkspaceId}/items?recursive=true";
                var itemsResponse = await httpClient.GetAsync(itemsUrl);
                var itemsJson = await itemsResponse.Content.ReadAsStringAsync();
                
                if (!itemsResponse.IsSuccessStatusCode)
                {
                    Console.WriteLine($"[SYNC] ERROR: Fabric Items API failed with {itemsResponse.StatusCode}. Content: {itemsJson}");
                }

                var itemsRoot = JsonConvert.DeserializeObject<dynamic>(itemsJson);
                var fabricItems = itemsRoot?.value;

                var localFolders = db.Folders.Where(f => f.WorkspaceId == localWorkspaceId).ToList();
                var localReports = db.Reports.Where(r => r.WorkspaceId == localWorkspaceId).ToList();
                var syncedReportIds = new HashSet<string>();

                // 3. Process every item found by Fabric
                if (fabricItems != null)
                {
                    foreach (var item in fabricItems)
                    {
                        string type = item.type?.ToString() ?? "";
                        string displayName = item.displayName?.ToString() ?? "Unknown";
                        string pbiId = item.id.ToString();
                        string? parentFolderId = item.parentFolderId?.ToString();
                        
                        if (processedPbiIds.Contains(pbiId)) 
                        {
                            Console.WriteLine($"[SYNC] Skipping already processed report: {pbiId}");
                            continue;
                        }
                        processedPbiIds.Add(pbiId);
                        syncedReportIds.Add(pbiId);

                        // Try to find datasetId from pbiReports list
                        var pbiMatch = pbiReports.FirstOrDefault(p => p.Id.ToString().Equals(pbiId, StringComparison.OrdinalIgnoreCase));
                        string? datasetId = pbiMatch?.DatasetId;

                        Console.WriteLine($"[SYNC] Found Fabric Item: {displayName} (Type: {type}) (Dataset: {datasetId ?? "NULL"})");

                        if (type != "Report" && type != "PaginatedReport") continue;

                        int? localFolderId = null;
                        if (!string.IsNullOrEmpty(parentFolderId) && localFolders != null)
                        {
                            var folderMatch = localFolders.FirstOrDefault(f => f.FabricFolderId != null && f.FabricFolderId.Equals(parentFolderId, StringComparison.OrdinalIgnoreCase));
                            localFolderId = folderMatch?.Id;
                            if (localFolderId.HasValue)
                                Console.WriteLine($"[SYNC] Report '{displayName}' belongs to Folder '{folderMatch?.Name}'");
                        }

                        string normalizedDisplayName = displayName.Replace(" ", "_").Replace("(", "").Replace(")", "");
                        
                        // 1. Try exact match by PowerBI ID
                        var existing = localReports.FirstOrDefault(r => (r.PowerBIReportId ?? "").Equals(pbiId, StringComparison.OrdinalIgnoreCase));

                        // 2. Fallback: Match by Name + Workspace (Only if the local report's current ID is missing or invalid)
                        if (existing == null)
                        {
                            // Check if the local record already has a valid ID that we haven't processed yet
                            var strayReport = localReports.FirstOrDefault(r => 
                                r.Name != null && r.Name.Equals(normalizedDisplayName, StringComparison.OrdinalIgnoreCase) && 
                                r.WorkspaceId == localWorkspaceId);

                            if (strayReport != null)
                            {
                                // CRITICAL: If the local record already has a valid ID that exists in Fabric,
                                // we MUST NOT overwrite it with a name-based duplicate.
                                bool currentIdStillExists = fabricItems != null && ((IEnumerable<dynamic>)fabricItems).Any(i => i.id != null && i.id.ToString().Equals(strayReport.PowerBIReportId ?? "", StringComparison.OrdinalIgnoreCase));
                                
                                if (!currentIdStillExists)
                                {
                                    Console.WriteLine($"[SYNC] ID Drift detected for '{strayReport.Name}'. Updating to latest ID: {pbiId}");
                                    strayReport.PowerBIReportId = pbiId;
                                    existing = strayReport;
                                }
                                else
                                {
                                    Console.WriteLine($"[SYNC] Skipping duplicate Fabric item '{displayName}' (ID: {pbiId}) because local report '{strayReport.Name}' is already synced to a valid ID ({strayReport.PowerBIReportId}).");
                                    continue; // Skip this fabric item, it's a redundant duplicate
                                }
                            }
                        }

                        if (existing == null)
                        {
                            Console.WriteLine($"[SERVICE] Adding new {type} to DB: {normalizedDisplayName} (Folder: {localFolderId})");
                            db.Reports.Add(new PowerBI.Models.Report
                            {
                                Name = normalizedDisplayName,
                                PowerBIReportId = pbiId,
                                PowerBIDatasetId = datasetId,
                                WorkspaceId = localWorkspaceId,
                                FolderId = localFolderId,
                                ReportType = type == "PaginatedReport" ? "RDL" : "PowerBI"
                            });
                        }
                        else
                        {
                            normalizedDisplayName = displayName.Replace(" ", "_").Replace("(", "").Replace(")", "");
                            if (existing.Name != normalizedDisplayName) existing.Name = normalizedDisplayName;
                            
                            // Only update folder if we found one in Fabric (prevents "un-moving" if API is slow)
                            if (localFolderId.HasValue) 
                            {
                                existing.FolderId = localFolderId;
                            }
                            
                            // CRITICAL FIX: Ensure DatasetId is saved for RDLs to avoid token failures
                            if (!string.IsNullOrEmpty(datasetId)) 
                            {
                                existing.PowerBIDatasetId = datasetId;
                            }
                            else if (existing.ReportType == "RDL")
                            {
                                // If DatasetId is missing from the report object, try to find the shadow dataset
                                var dsMatch = pbiReports.FirstOrDefault(p => p.Name == displayName || p.Name == displayName + ".rdl");
                                if (dsMatch != null) existing.PowerBIDatasetId = dsMatch.DatasetId;
                            }

                            existing.ReportType = type == "PaginatedReport" ? "RDL" : "PowerBI";
                        }
                    }
                }

                // 4. Cleanup: Remove local reports that no longer exist in Fabric
                if (syncedReportIds.Any())
                {
                    foreach (var localRep in localReports)
                    {
                        if (localRep.PowerBIReportId != null && !syncedReportIds.Contains(localRep.PowerBIReportId))
                        {
                            Console.WriteLine($"[SERVICE] Removing report from DB (deleted in PBI): {localRep.Name}");
                            db.Reports.Remove(localRep);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("[SYNC] WARNING: No reports found in Fabric. Skipping cleanup to prevent accidental data loss.");
                }

                await db.SaveChangesAsync();

                // 5. AUTO-CLEANUP: If we have multiple reports with the same name in Fabric, cleanup the duplicates
                try {
                    if (fabricItems != null)
                    {
                        var itemsList = ((IEnumerable<dynamic>)fabricItems).ToList();
                        var duplicateGroups = itemsList
                            .Where(i => (string)(i.type?.ToString() ?? "") == "PaginatedReport" || (string)(i.type?.ToString() ?? "") == "Report")
                            .GroupBy(i => (string)(i.displayName?.ToString() ?? "Unknown"))
                            .Where(g => g.Count() > 1);

                        foreach (var group in duplicateGroups)
                        {
                            string reportName = group.Key; // Explicitly cast to string
                            string normalizedName = reportName.Replace(" ", "_").Replace("(", "").Replace(")", "");

                            // Keep the one that is currently in our DB, delete others from Fabric
                            var masterRep = db.Reports.FirstOrDefault(r => r.WorkspaceId == localWorkspaceId && r.Name == normalizedName);
                            
                            if (masterRep != null)
                            {
                                foreach (var item in group)
                                {
                                    string pbiId = item.id.ToString();
                                    if (pbiId != masterRep.PowerBIReportId)
                                    {
                                        Console.WriteLine($"[AUTO-CLEANUP] Deleting redundant duplicate from Fabric: {reportName} (ID: {pbiId})");
                                        await DeleteReport(pbiWorkspaceId, Guid.Parse(pbiId));
                                        
                                        // CRITICAL: Also remove from local DB so user doesn't click a dead link
                                        var staleRep = db.Reports.FirstOrDefault(r => r.PowerBIReportId == pbiId);
                                        if (staleRep != null) db.Reports.Remove(staleRep);
                                    }
                                }
                            }
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine($"[AUTO-CLEANUP] Warning: {ex.Message}");
                }
            } // End of Workspace Loop
        } // End of SyncReports Method

        public async Task<(Import Import, string? RdlContent)> UploadReport(int localWorkspaceId, Guid pbiWorkspaceId, string name, Stream stream, AppDbContext db, string? folderName = null, List<string>? customParams = null)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] REQUEST: Upload '{name}' (Size: {stream.Length} bytes) (Custom Params: {customParams?.Count ?? 0})");
                var client = await GetClient();

                string extension = Path.GetExtension(name).ToLower();
                Console.WriteLine($"[SERVICE] Detected Extension: '{extension}'");

                if (extension == ".pbit")
                {
                    throw new Exception("Power BI Templates (.pbit) are not supported for direct upload. Please save your file as a .pbix and try again.");
                }

                var isRdl = name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"[SERVICE] RDLLLLLLL.............: {isRdl}");

                var datasetName = Path.GetFileNameWithoutExtension(name);
                
                // For RDL, the display name often NEEDS the .rdl extension to be accepted as a Paginated Report
                var finalDisplayName = isRdl ? name : Path.GetFileNameWithoutExtension(name);

                var targetFolder = string.IsNullOrEmpty(folderName) ? "Automated Reports" : folderName;

                // --- RDL OVERWRITE WORKAROUND ---
                // Power BI API does not support CreateOrOverwrite for RDL. 
                // We must delete the old one first if it exists to simulate an update.
                if (isRdl)
                {
                    Console.WriteLine($"[SERVICE] RDL detected. Checking for existing report to replace...");
                    var existingReports = await client.Reports.GetReportsInGroupAsync(pbiWorkspaceId);
                    var oldReport = existingReports.Value.FirstOrDefault(r => 
                        string.Equals(r.Name, finalDisplayName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(Path.GetFileNameWithoutExtension(r.Name), finalDisplayName, StringComparison.OrdinalIgnoreCase));
                    
                    if (oldReport != null)
                    {
                        Console.WriteLine($"[SERVICE] Found existing RDL '{oldReport.Name}' (ID: {oldReport.Id}). Deleting for replacement...");
                        await client.Reports.DeleteReportInGroupAsync(pbiWorkspaceId, oldReport.Id);
                        Console.WriteLine("[SERVICE] Old RDL deleted. Proceeding with clean upload.");
                    }
                }

                // Paginated Reports (.rdl) ONLY support 'Abort' for nameConflict. 
                // We handle 'Overwrite' manually by deleting the existing report first (lines 542-556).
                var conflictMode = isRdl ? ImportConflictHandlerMode.Abort : ImportConflictHandlerMode.CreateOrOverwrite;

                Console.WriteLine($"[SERVICE] DESTINATION: Workspace {pbiWorkspaceId}");
                Console.WriteLine($"[SERVICE] ACTION: Uploading {name} as '{finalDisplayName}' (Mode: {conflictMode})");

                int? folderId = null;
                try 
                {
                    folderId = await GetOrCreateFolder(localWorkspaceId, pbiWorkspaceId, targetFolder, db);
                    Console.WriteLine($"[SERVICE] Target Folder ID: {folderId}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SERVICE] Folder system not ready or unsupported: {ex.Message}. Falling back to root upload.");
                }

                // SANITIZE DISPLAY NAME (Remove spaces/brackets which can cause BadRequest in some API paths)
                finalDisplayName = finalDisplayName.Replace(" ", "_").Replace("(", "").Replace(")", "");
                Console.WriteLine($"[SERVICE] Sanitized Display Name: '{finalDisplayName}'");

                Import import;
                Stream uploadStream = stream;
                string? updatedXml = null;

                // --- DB PERSISTENCE & CONNECTION INJECTION ---
                if (isRdl)
                {
                    Console.WriteLine($">>>> [RDL-UPLOAD] Processing '{name}' for SQL Persistence...");
                    var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                    if (!Directory.Exists(uploadsDir)) Directory.CreateDirectory(uploadsDir);
                    var filePath = Path.Combine(uploadsDir, finalDisplayName);
                    
                    // 1. Load XML for manipulation
                    XDocument xDoc;
                    using (var tempMs = new MemoryStream())
                    {
                        await stream.CopyToAsync(tempMs);
                        tempMs.Position = 0;
                        xDoc = XDocument.Load(tempMs);
                        stream.Position = 0; // Reset original stream
                    }
                    var ns = xDoc.Root.GetDefaultNamespace();
                    Console.WriteLine(">>>> [RDL-UPLOAD] XML loaded into memory.");

                    // 1.5 Strip Parameter Layout (Prevent rsInvalidParameterLayoutCellDefNotEqualsParameterCount)
                    var layout = xDoc.Descendants(ns + "ReportParametersLayout").FirstOrDefault();
                    if (layout != null)
                    {
                        Console.WriteLine(">>>> [RDL-UPLOAD] Stripping stale Parameter Layout grid.");
                        layout.Remove();
                    }

                    // 2. Inject Connection String (Convert Static to SQL)
                    string? defaultConn = _config.GetConnectionString("DefaultConnection");
                    if (!string.IsNullOrEmpty(defaultConn))
                    {
                        Console.WriteLine(">>>> [RDL-UPLOAD] Scanning for DataSources to inject ConnectionString...");
                        var dataSources = xDoc.Descendants(ns + "DataSource");
                        foreach (var ds in dataSources)
                        {
                            var connProps = ds.Element(ns + "ConnectionProperties");
                            if (connProps != null)
                            {
                                // --- STATIC DATA PROTECTION ---
                                // Check if any dataset using this datasource contains XML/Static data.
                                // If it does, forcing 'SQL' provider will break the report.
                                string dsName = ds.Attribute("Name")?.Value ?? "";
                                bool hasStaticData = xDoc.Descendants(ns + "DataSet")
                                    .Where(d => d.Element(ns + "Query")?.Element(ns + "DataSourceName")?.Value == dsName)
                                    .Any(d => d.Element(ns + "Query")?.Element(ns + "CommandText")?.Value?.TrimStart().StartsWith("<Query", StringComparison.OrdinalIgnoreCase) == true);

                                if (hasStaticData)
                                {
                                    Console.WriteLine($">>>> [RDL-UPLOAD] SKIP: DataSource '{dsName}' uses Static XML Data. Bypassing SQL Injection.");
                                    continue;
                                }

                                var provider = connProps.Element(ns + "DataProvider");
                                if (provider == null) 
                                { 
                                    provider = new XElement(ns + "DataProvider"); 
                                    connProps.AddFirst(provider); 
                                }
                                
                                // FORCE: Use 'SQL' (SQL Server) extension for local SQL connections.
                                if (provider.Value != "SQL")
                                {
                                    Console.WriteLine($">>>> [RDL-UPLOAD] Updating DataProvider for '{dsName}' from '{provider.Value}' to 'SQL'.");
                                    provider.Value = "SQL";
                                }

                                var connectString = connProps.Element(ns + "ConnectString");
                                if (connectString == null) { connectString = new XElement(ns + "ConnectString"); connProps.Add(connectString); }
                                connectString.Value = defaultConn;
                                
                                Console.WriteLine($">>>> [RDL-UPLOAD] ConnectionString injected into DataSource '{dsName}'.");
                            }
                        }
                    }

                    // 3. Save modified XML to string for DB storage
                    updatedXml = xDoc.ToString();
                    Console.WriteLine($">>>> [RDL-UPLOAD] RDL Content prepared (Size: {updatedXml.Length} chars).");
                    
                    // 4. Save locally for legacy discovery tools
                    var utf8WithBom = new System.Text.UTF8Encoding(true);
                    using (var sw = new StreamWriter(filePath, false, utf8WithBom)) { xDoc.Save(sw); }
                    Console.WriteLine($">>>> [RDL-UPLOAD] Temporary RDL file saved to uploads folder.");
                    
                    // 5. Update upload stream
                    byte[] fileBytes = utf8WithBom.GetBytes(updatedXml);
                    uploadStream = new MemoryStream(fileBytes);
                }
                
                // Step 1: Upload to Root
                Console.WriteLine($"[SERVICE] Sending file to Power BI: {finalDisplayName} (Size: {uploadStream.Length} bytes)");
                
                try 
                {
                    import = await client.Imports.PostImportWithFileAsync(
                        pbiWorkspaceId,
                        uploadStream,
                        datasetDisplayName: finalDisplayName, 
                        nameConflict: conflictMode
                    );
                    Console.WriteLine($"[SERVICE] Upload request accepted. Import ID: {import?.Id}");
                }
                catch (HttpOperationException ex)
                {
                    Console.WriteLine($"[SERVICE] CRITICAL UPLOAD ERROR (HTTP): {ex.Response.Content}");
                    throw new Exception($"Power BI Upload Failed: {ex.Response.ReasonPhrase}. Details: {ex.Response.Content}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SERVICE] CRITICAL UPLOAD ERROR: {ex.Message}");
                    if (ex.InnerException != null) Console.WriteLine($"[SERVICE] INNER ERROR: {ex.InnerException.Message}");
                    throw;
                }

                // Step 2: Poll for completion
                Console.WriteLine($"[SERVICE] Import ID: {import.Id}. Monitoring status...");
                int attempts = 0;
                while (import.ImportState != "Succeeded" && import.ImportState != "Failed" && attempts < 20)
                {
                    await Task.Delay(2000);
                    import = await client.Imports.GetImportInGroupAsync(pbiWorkspaceId, import.Id);
                    Console.WriteLine($"[SERVICE] Status: {import.ImportState} (Attempt {attempts+1})");
                    attempts++;
                }

                if (import.ImportState != "Succeeded")
                {
                    throw new Exception($"Power BI Import failed with state: {import.ImportState}");
                }

                // Step 3: Move to Folder
                if (folderId.HasValue)
                {
                    var folder = await db.Folders.FindAsync(folderId.Value);
                    if (folder != null && !string.IsNullOrEmpty(folder.FabricFolderId))
                    {
                        Console.WriteLine($"[SERVICE] Preparing to move report to folder '{folder.Name}'...");
                        Guid? reportToMove = null;
                        
                        // Wait and search loop
                        for (int searchAttempt = 1; searchAttempt <= 4; searchAttempt++)
                        {
                            Console.WriteLine($"[SERVICE] Discovery Attempt {searchAttempt} for '{finalDisplayName}'...");
                            
                            // 1. Try finding in the import result again (sometimes it populates late)
                            if (import.Reports != null && import.Reports.Any())
                            {
                                reportToMove = import.Reports.First().Id;
                            }

                            // 2. Scan workspace aggressively
                            if (!reportToMove.HasValue)
                            {
                                var allReports = await client.Reports.GetReportsInGroupAsync(pbiWorkspaceId);
                                Console.WriteLine($"[SERVICE] Workspace Scan: Found {allReports.Value.Count} total reports.");
                                
                                foreach (var r in allReports.Value)
                                {
                                    // Match by name (exact or without extension)
                                    if (string.Equals(r.Name, finalDisplayName, StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(Path.GetFileNameWithoutExtension(r.Name), finalDisplayName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        reportToMove = r.Id;
                                        break;
                                    }
                                }
                            }

                            if (reportToMove.HasValue) break;

                            Console.WriteLine("[SERVICE] Report not found yet. Waiting 4 seconds...");
                            await Task.Delay(4000);
                            import = await client.Imports.GetImportInGroupAsync(pbiWorkspaceId, import.Id);
                        }

                        if (reportToMove.HasValue)
                        {
                            Console.WriteLine($"[SERVICE] SUCCESS: Found Report ID {reportToMove.Value}. Moving to Fabric Folder {folder.FabricFolderId}...");
                            await MoveItemToFolder(pbiWorkspaceId, reportToMove.Value, Guid.Parse(folder.FabricFolderId));
                            Console.WriteLine("[SERVICE] MOVE COMPLETED SUCCESSFULLY.");
                            
                            // Hotfix: Ensure credentials/datasources are synced after move
                            await PatchReportCredentials(pbiWorkspaceId, reportToMove.Value);

                            // Attach the discovered ID to the import object for the caller (hotfix) to use
                            import.Reports = new List<Microsoft.PowerBI.Api.Models.Report> { new Microsoft.PowerBI.Api.Models.Report { Id = reportToMove.Value } };
                        }
                        else
                        {
                            Console.WriteLine("[SERVICE] ERROR: Could not find report in workspace after multiple attempts. It may be stuck in Root.");
                        }
                    }
                }

                return (import, isRdl ? updatedXml : null);
            }
            catch (Exception ex)
            {
                throw new Exception($"Upload failed: {ex.Message}");
            }
        }

        // --- REMOVED LEGACY CLONING LOGIC (DEPRECATED FOR SINGLE REPORT ARCHITECTURE) ---



        public async Task<Stream> ExportReportAsStream(Guid workspaceId, Guid reportId, List<ExportFilter>? filters = null, string reportType = "PowerBI", AppDbContext? db = null)
        {
            Console.WriteLine($"[SERVICE] Exporting {reportType} Report {reportId} to PDF...");
            var client = await GetClient();

            var exportRequest = new ExportReportRequest { Format = FileFormat.PDF };

            if (reportType == "RDL")
            {
                // RDL (Paginated) reports use ParameterValues
                var rdlParams = new List<ParameterValue>();
                if (filters != null)
                {
                    foreach (var f in filters)
                    {
                        var val = f.Filter.Split("eq").Last().Trim().Trim('\'');
                        var paramName = f.Filter.Split('/').Last().Split(' ').First();
                        
                        rdlParams.Add(new ParameterValue { Name = paramName, Value = val });
                        Console.WriteLine($"[SERVICE] RDL Export Param: {paramName} = {val}");
                    }
                }

                // ONLY include background TenantId if the report was discovered to have it
                if (db != null)
                {
                    // Find the local report record first
                    var localRep = db.Reports.FirstOrDefault(r => r.PowerBIReportId == reportId.ToString());
                    if (localRep != null && db.ReportFilters != null)
                    {
                        var hasSecurityParam = db.ReportFilters.Any(f => f.ReportId == localRep.Id && f.ColumnName == "TenantId");
                        if (hasSecurityParam && !rdlParams.Any(p => p.Name == "TenantId"))
                        {
                            Console.WriteLine("[SERVICE] Injecting background TenantId into PDF export (Security requirement met).");
                            rdlParams.Add(new ParameterValue { Name = "TenantId", Value = "PROD-TENANT-001" });
                        }
                        else if (!hasSecurityParam)
                        {
                            Console.WriteLine("[SERVICE] Skipping background TenantId injection (Parameter not found in report schema).");
                        }
                    }
                }

                exportRequest.PaginatedReportConfiguration = new PaginatedReportExportConfiguration
                {
                    ParameterValues = rdlParams
                };
            }
            else
            {
                // Standard Power BI reports use ReportLevelFilters
                exportRequest.PowerBIReportConfiguration = new Microsoft.PowerBI.Api.Models.PowerBIReportExportConfiguration
                {
                    ReportLevelFilters = filters
                };
            }
            
            try {
                var export = await client.Reports.ExportToFileInGroupAsync(workspaceId, reportId, exportRequest);

            Export status;
            do
            {
                await Task.Delay(3000);
                status = await client.Reports.GetExportToFileStatusInGroupAsync(workspaceId, reportId, export.Id);
                Console.WriteLine($"[SERVICE] Export Status: {status.Status}");
            } while (status.Status != ExportState.Succeeded && status.Status != ExportState.Failed);

            if (status.Status == ExportState.Failed) throw new Exception("Power BI Export Failed.");

                return await client.Reports.GetFileOfExportToFileInGroupAsync(workspaceId, reportId, export.Id);
            }
            catch (Microsoft.Rest.HttpOperationException ex) when (ex.Response.StatusCode == System.Net.HttpStatusCode.Forbidden) {
                throw new Exception("Export Forbidden. Please ensure: 1. Service Principal is enabled for 'Export reports as PDF' in Power BI Admin Portal. 2. The Workspace has a Premium/Fabric capacity.");
            }
        }

        public async Task<ReportEmbedConfig> GetEmbedConfig(Guid workspaceId, Guid reportId, AppDbContext db, List<ParameterValue>? parameterValues = null)
        {
            Console.WriteLine($"[SERVICE] Fetching Embed Config for Report {reportId}");
            
            // 1. Fetch from local DB to get cached DatasetId
            var localReport = await db.Reports.FirstOrDefaultAsync(r => r.PowerBIReportId == reportId.ToString());
            string? targetDatasetId = localReport?.PowerBIDatasetId;

            // DIAGNOSTIC: Check datasource status before embedding
            await PatchReportCredentials(workspaceId, reportId);
            
            var client = await GetClient();

            try {
                var report = await client.Reports.GetReportInGroupAsync(workspaceId, reportId);
                
                // DIAGNOSTIC: Check Capacity Assignment
                var group = await client.Groups.GetGroupsAsync(filter: $"id eq '{workspaceId}'");
                var workspaceInfo = group.Value.FirstOrDefault();
                string capacityStr = workspaceInfo?.CapacityId?.ToString() ?? "NONE";
                Console.WriteLine($">>>> [CAPACITY] Workspace: {workspaceInfo?.Name}, CapacityId: {capacityStr}");
                
                if (capacityStr == "NONE" || workspaceInfo?.CapacityId == Guid.Empty)
                {
                    Console.WriteLine(">>>> [CAPACITY] WARNING: Workspace is NOT assigned to a capacity. RDLs might not render for Service Principal.");
                }

                var accessToken = await _auth.GetAccessToken();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                
                // --- RDL DATA ACCESS FIX (SUPER-DISCOVERY) ---
                bool isRdl = report.ReportType == "RDL" || report.ReportType == "PaginatedReport";

                if (!isRdl && string.IsNullOrEmpty(targetDatasetId))
                {
                    targetDatasetId = report.DatasetId;
                    
                    if (string.IsNullOrEmpty(targetDatasetId))
                    {
                        Console.WriteLine($">>>> [TOKEN] DatasetId null for '{report.Name}'. Scanning all Workspace Datasets...");
                        var allDatasets = await client.Datasets.GetDatasetsInGroupAsync(workspaceId);
                        
                        foreach (var ds in allDatasets.Value)
                        {
                            Console.WriteLine($">>>> [TOKEN]   Scanning Dataset: '{ds.Name}' (ID: {ds.Id})");
                            try 
                            {
                                var dsSources = await client.Datasets.GetDatasourcesInGroupAsync(workspaceId, ds.Id);
                                foreach(var source in dsSources.Value)
                                {
                                    string server = source.ConnectionDetails?.Server ?? "UNKNOWN";
                                    string dbName = source.ConnectionDetails?.Database ?? "UNKNOWN";
                                    Console.WriteLine($">>>> [TOKEN]   Checking Dataset '{ds.Name}': Server='{server}', DB='{dbName}'");

                                    if (server.Contains("zsdeskt8ibnl5", StringComparison.OrdinalIgnoreCase) || 
                                        server.Contains("localhost", StringComparison.OrdinalIgnoreCase))
                                    {
                                        targetDatasetId = ds.Id;
                                        Console.WriteLine($">>>> [TOKEN] FOUND MATCH! Mapping Report to Dataset ID: {targetDatasetId}");
                                        break;
                                    }
                                }
                            } catch { }
                            if (!string.IsNullOrEmpty(targetDatasetId)) break;
                        }

                        if (string.IsNullOrEmpty(targetDatasetId))
                        {
                            Console.WriteLine(">>>> [TOKEN] DNA Scan failed. Trying Name Match...");
                            var allDatasetsList = allDatasets.Value.ToList();
                            var nameMatch = allDatasetsList.FirstOrDefault(d => 
                                d.Name.Contains(report.Name, StringComparison.OrdinalIgnoreCase) || 
                                report.Name.Contains(d.Name, StringComparison.OrdinalIgnoreCase));
                                
                            if (nameMatch != null) 
                            {
                                targetDatasetId = nameMatch.Id;
                                Console.WriteLine($">>>> [TOKEN] NAME MATCH! Using Dataset: {nameMatch.Name} ({targetDatasetId})");
                            }
                        }

                        if (string.IsNullOrEmpty(targetDatasetId))
                        {
                            Console.WriteLine(">>>> [TOKEN] WARNING: No shadow dataset found. Using ReportId as DatasetId fallback...");
                            targetDatasetId = report.Id.ToString();
                        }
                    }
                }
                else if (isRdl)
                {
                    Console.WriteLine($">>>> [TOKEN] RDL Detected: Proceeding with cautious dataset discovery.");
                    // We don't null it here anymore; we let the scan logic above (DNA/Name match) try to find the shadow dataset.
                }


                // --- AZURE-IT-SPEC: RAW REST TOKEN GENERATION ---
                Console.WriteLine($">>>> [TOKEN] Generating Token for Report: {report.Id}");
                
                // NOTE: Datasource identities and manual binding are bypassed here for RDLs
                // because 'PatchReportCredentials' (called above) already established the Gateway handshake.

                // --- AZURE-IT-SPEC: MODERN V2 GENERATE TOKEN ---
                var v2TokenRequest = new Dictionary<string, object>
                {
                    { "reports", new[] { new { id = report.Id.ToString() } } },
                    { "targetWorkspaces", new[] { new { id = workspaceId.ToString() } } },
                    { "xmlaPermissions", "ReadOnly" }
                };

                // Add datasets if found (helps authorize RDL data sources)
                if (!string.IsNullOrEmpty(targetDatasetId))
                {
                    v2TokenRequest["datasets"] = new[] { new { id = targetDatasetId } };
                }
                else if (report.ReportType != "RDL" && report.ReportType != "PaginatedReport")
                {
                    // Fallback for standard reports using SDK property
                    v2TokenRequest["datasets"] = new[] { new { id = report.DatasetId } };
                }

                string tokenUrl = "https://api.powerbi.com/v1.0/myorg/GenerateToken";
                var tokenResponse = await httpClient.PostAsync(tokenUrl, new StringContent(JsonConvert.SerializeObject(v2TokenRequest), Encoding.UTF8, "application/json"));
                
                if (!tokenResponse.IsSuccessStatusCode)
                {
                    var error = await tokenResponse.Content.ReadAsStringAsync();
                    Console.WriteLine($">>>> [TOKEN] FAILED: {error}");
                    throw new Exception($"Token API failed: {error}");
                }

                var tokenResult = JsonConvert.DeserializeObject<dynamic>(await tokenResponse.Content.ReadAsStringAsync());
                string embedToken = tokenResult?.token;
                Console.WriteLine("[SERVICE] Embed Token generated successfully.");

                var embedUrl = report.EmbedUrl;
                
                // Convert to RDL embed if needed (CRITICAL for RDL rendering)
                if (report.ReportType == "PaginatedReport" && !embedUrl.Contains("rdlEmbed"))
                {
                    embedUrl = embedUrl.Replace("reportEmbed", "rdlEmbed");
                }

                // --- SYNC PARAMETERS INTO URL (FOR STABILITY) ---
                if (isRdl || (parameterValues != null && parameterValues.Any()))
                {
                    var rdlUrlParams = new List<string>();
                    
                    if (isRdl)
                    {
                        rdlUrlParams.Add("rdl:parameterPanel=collapsed");
                    }

                    if (parameterValues != null && parameterValues.Any())
                    {
                        foreach (var p in parameterValues)
                        {
                            rdlUrlParams.Add($"rp:{Uri.EscapeDataString(p.Name)}={Uri.EscapeDataString(p.Value ?? "")}");
                        }
                    }

                    if (rdlUrlParams.Any())
                    {
                        string separator = embedUrl.Contains("?") ? "&" : "?";
                        embedUrl += separator + string.Join("&", rdlUrlParams);
                    }
                    
                    Console.WriteLine($">>>> [TRACE] FINAL EMBED URL: {embedUrl}");
                }

                return new ReportEmbedConfig
                {
                    ReportId = report.Id.ToString(),
                    DatasetId = report.DatasetId,
                    EmbedUrl = embedUrl,
                    EmbedToken = embedToken,
                    ReportName = report.Name,
                    ReportType = report.ReportType,
                    LocalReportId = localReport?.Id ?? 0,
                    WorkspaceId = workspaceId.ToString(),
                    TenantId = _config["PowerBI:TenantId"] ?? "UNKNOWN",
                    Parameters = parameterValues?.ToDictionary(p => p.Name ?? "", p => p.Value ?? "") ?? new()
                };
            }
            catch (Microsoft.Rest.HttpOperationException ex) when (ex.Response.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                throw new Exception("The report was not found in Power BI. It may have been deleted or moved. Please try syncing your workspace again.");
            }
        }


        public async Task<bool> DiscoverReportFilters(int reportId, Guid? datasetId, AppDbContext db, int userId)
        {
            bool modified = false;
            var report = await db.Reports.FindAsync(reportId);
            if (report == null) return false;

            var workspace = await db.Workspaces.FindAsync(report.WorkspaceId);
            if (workspace == null || string.IsNullOrEmpty(workspace.PowerBIWorkspaceId))
                throw new Exception("Workspace not found or not synced.");

            Console.WriteLine($"[SCHEMA-DISCOVERY] MODE: {report.ReportType} (USER: {userId})");
            Console.WriteLine($"[SCHEMA-DISCOVERY] Report Name: {report.Name}");

            var existingFilters = db.ReportFilters
                .Where(f => f.ReportId == reportId && f.UserId == userId)
                .ToList();

            var processedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int count = 0;

            if (report.ReportType == "RDL")
            {
                try
                {
                    Console.WriteLine($">>>> [FLOW] RDL: Attempting API discovery for report '{report.Name}'...");
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspace.PowerBIWorkspaceId}/reports/{report.PowerBIReportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($">>>> [FLOW] RDL: API discovery successful.");
                        var rawJson = await response.Content.ReadAsStringAsync();
                        var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                        var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                string? paramName = param["name"]?.ToString();
                                if (string.IsNullOrEmpty(paramName)) continue;
                                Console.WriteLine($">>>> [FLOW] RDL: Identified Parameter (API): {paramName}");

                                var existing = existingFilters.FirstOrDefault(f => f.TableName == "RDL_PARAMETER" && f.ColumnName == paramName);
                                if (existing == null)
                                {
                                    existing = new ReportFilter
                                    {
                                        ReportId = reportId,
                                        TableName = "RDL_PARAMETER",
                                        ColumnName = paramName,
                                        DisplayName = paramName,
                                        IsActive = true,
                                        UserId = userId,
                                        IsCustom = false
                                    };
                                    db.ReportFilters.Add(existing);
                                }
                                else if (!existing.IsCustom)
                                {
                                    existing.IsActive = true;
                                }
                                processedKeys.Add($"RDL_PARAMETER|{paramName}");
                                count++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($">>>> [FLOW] RDL: API returned {response.StatusCode}. Falling back to Local Parsing...");
                        string? rdlXml = report.RdlContent;
                        XDocument xDoc;

                        if (!string.IsNullOrEmpty(rdlXml))
                        {
                            xDoc = XDocument.Parse(rdlXml);
                        }
                        else
                        {
                            var rdlName = report.Name ?? "Unknown";
                            var fName = rdlName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? rdlName : $"{rdlName}.rdl";
                            var fPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fName);
                            if (File.Exists(fPath)) xDoc = XDocument.Load(fPath);
                            else return false;
                        }

                        var ns = xDoc.Root.GetDefaultNamespace();
                        var rdlParamsList = xDoc.Descendants(ns + "ReportParameter");

                        foreach (var p in rdlParamsList)
                        {
                            string? paramName = p.Attribute("Name")?.Value;
                            if (string.IsNullOrEmpty(paramName)) continue;
                            
                            var existing = existingFilters.FirstOrDefault(f => f.TableName == "RDL_PARAMETER" && f.ColumnName == paramName);
                            if (existing == null)
                            {
                                existing = new ReportFilter { ReportId = reportId, TableName = "RDL_PARAMETER", ColumnName = paramName, DisplayName = paramName, IsActive = true, UserId = userId, IsCustom = false };
                                db.ReportFilters.Add(existing);
                            }
                            processedKeys.Add($"RDL_PARAMETER|{paramName}");
                            count++;
                        }

                        // SELF-HEALING & UNIVERSAL SYNC
                        // We force injection for all active custom filters to ensure they are applied to ALL datasets 
                        // (handles the case where a filter was previously injected but missed some datasets).
                        var activeCustoms = existingFilters.Where(f => f.IsCustom && f.IsActive).ToList();
                        foreach (var m in activeCustoms)
                        {
                            var injectRes = await InjectFilterToRdl(reportId, m.ColumnName, db);
                            if (injectRes.Success)
                            {
                                if (!processedKeys.Contains($"{m.TableName}|{m.ColumnName}"))
                                    processedKeys.Add($"{m.TableName}|{m.ColumnName}");
                                modified = true;
                            }
                        }
                    }
                }
                catch (Exception ex) { Console.WriteLine($"[SCHEMA-DISCOVERY] RDL Error: {ex.Message}"); }
            }
            else
            {
                if (!datasetId.HasValue) throw new Exception("Dataset ID is missing for Power BI report.");
                var token = await _auth.GetAccessToken();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var url = $"https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/tables";
                var response = await httpClient.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var rawJson = await response.Content.ReadAsStringAsync();
                    var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                    var tables = root["value"] as Newtonsoft.Json.Linq.JArray;
                    if (tables != null)
                    {
                        foreach (var tbl in tables)
                        {
                            string? tableName = tbl["name"]?.ToString();
                            if (string.IsNullOrEmpty(tableName)) continue;
                            var cols = tbl["columns"] as Newtonsoft.Json.Linq.JArray;
                            if (cols != null)
                            {
                                foreach (var col in cols)
                                {
                                    string? colName = col["name"]?.ToString();
                                    if (string.IsNullOrEmpty(colName)) continue;
                                    var existing = existingFilters.FirstOrDefault(f => f.TableName == tableName && f.ColumnName == colName);
                                    if (existing == null)
                                    {
                                        existing = new ReportFilter { ReportId = reportId, TableName = tableName, ColumnName = colName, DisplayName = colName, IsActive = count < 6, UserId = userId };
                                        db.ReportFilters.Add(existing);
                                    }
                                    processedKeys.Add($"{tableName}|{colName}");
                                    count++;
                                }
                            }
                        }
                    }
                }
            }

            foreach (var f in existingFilters)
            {
                if (!processedKeys.Contains($"{f.TableName}|{f.ColumnName}") && !f.IsCustom)
                    db.ReportFilters.Remove(f);
            }

            if (count == 0) throw new Exception("No queryable fields or parameters discovered.");
            await db.SaveChangesAsync();
            return modified;
        }




        public async Task<List<string>> GetColumnValues(Guid? datasetId, string tableName, string columnName, Guid? reportId = null, Guid? workspaceId = null, AppDbContext? db = null, string reportType = "PowerBI")
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
                throw new Exception("[DYNAMIC] Error: Missing Table or Column name.");

            if (tableName == "RDL_PARAMETER")
            {
                if (!reportId.HasValue || !workspaceId.HasValue) throw new Exception("Report/Workspace ID required for RDL values.");
                
                try
                {
                    Console.WriteLine($">>>> [TERMINAL-LOG] Fetching values for Parameter: '{columnName}' (Report: {reportId})");
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspaceId}/reports/{reportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        url = $"https://api.powerbi.com/v1.0/myorg/reports/{reportId}/parameters";
                        response = await httpClient.GetAsync(url);
                    }

                    if (response.IsSuccessStatusCode)
                    {
                        var rawJson = await response.Content.ReadAsStringAsync();
                        var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                        var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                if (param["name"]?.ToString().Equals(columnName, StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    var suggestedValues = param["suggestedValues"] as Newtonsoft.Json.Linq.JArray;
                                    if (suggestedValues != null)
                                    {
                                        return suggestedValues.Select(v => v.ToString()).ToList();
                                    }
                                }
                            }
                        }
                    }

                    // FALLBACK: Parse XML for ValidValues
                    var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                    string? specificFile = null;
                    if (reportId.HasValue && db != null)
                    {
                        var reportRecord = await db.Reports.FirstOrDefaultAsync(r => r.PowerBIReportId == reportId.Value.ToString());
                        if (reportRecord != null)
                        {
                            string reportName = reportRecord.Name ?? "UnknownReport";
                            var fileName = reportName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? reportName : $"{reportName}.rdl";
                            var path = Path.Combine(uploadsDir, fileName);
                            if (File.Exists(path)) specificFile = path;
                        }
                    }

                    var filesToScan = (specificFile != null) ? new[] { specificFile } : Directory.GetFiles(uploadsDir, "*.rdl");
                    
                    foreach (var file in filesToScan)
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(file);
                        XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
                        if (xmlDoc.DocumentElement != null)
                            nsMgr.AddNamespace("rdl", xmlDoc.DocumentElement.NamespaceURI);

                        var paramNode = xmlDoc.SelectSingleNode($"//rdl:ReportParameter[@Name='{columnName}']", nsMgr);
                        if (paramNode != null)
                        {
                            // 1. Check for static valid values
                            var validValues = paramNode.SelectNodes("rdl:ValidValues/rdl:ParameterValues/rdl:ParameterValue", nsMgr);
                            if (validValues != null && validValues.Count > 0)
                            {
                                var xmlResults = new List<string>();
                                foreach (XmlNode val in validValues)
                                {
                                    string? label = val.SelectSingleNode("rdl:Label", nsMgr)?.InnerText;
                                    string? value = val.SelectSingleNode("rdl:Value", nsMgr)?.InnerText;
                                    xmlResults.Add(label ?? value ?? "");
                                }
                                return xmlResults.Where(s => !string.IsNullOrEmpty(s)).Distinct().ToList();
                            }

                            // 2. Check for DataSetReference (Dynamic/Embedded Values)
                            var dsRef = paramNode.SelectSingleNode("rdl:ValidValues/rdl:DataSetReference", nsMgr);
                            if (dsRef != null)
                            {
                                string? dsName = dsRef.SelectSingleNode("rdl:DataSetName", nsMgr)?.InnerText;
                                string? valueField = dsRef.SelectSingleNode("rdl:ValueField", nsMgr)?.InnerText;
                                
                                if (!string.IsNullOrEmpty(dsName) && !string.IsNullOrEmpty(valueField))
                                {
                                    var dsNode = xmlDoc.SelectSingleNode($"//rdl:DataSet[@Name='{dsName}']", nsMgr);
                                    var commandText = dsNode?.SelectSingleNode("rdl:Query/rdl:CommandText", nsMgr)?.InnerText;
                                    
                                    if (!string.IsNullOrEmpty(commandText) && commandText.Contains("<XmlData>"))
                                    {
                                        try
                                        {
                                            // Handle escaped XML within CommandText
                                            string rawXml = commandText.Trim();
                                            if (rawXml.StartsWith("<Query"))
                                            {
                                                XmlDocument innerXml = new XmlDocument();
                                                innerXml.LoadXml(rawXml);
                                                var dataNodes = innerXml.SelectNodes($"//Row/{valueField}");
                                                if (dataNodes != null)
                                                {
                                                    var dsResults = new List<string>();
                                                    foreach (XmlNode node in dataNodes) dsResults.Add(node.InnerText);
                                                    return dsResults.Where(s => !string.IsNullOrEmpty(s)).Distinct().ToList();
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($">>>> [TERMINAL-LOG] Failed to parse embedded XmlData: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    return new List<string>();
                }
                catch (Exception ex)
                { 
                    Console.WriteLine($">>>> [TERMINAL-LOG] RDL Value Error: {ex.Message}");
                    return new List<string>(); 
                }
            }
            if (reportType == "RDL" || reportType == "PaginatedReport")
            {
                Console.WriteLine($">>>> [FLOW] Skipping DAX probe for {reportType} report field '{columnName}'. RDLs only support parameter-based value discovery.");
                return new List<string>(); // RDLs don't support DAX queries on datasets via this API
            }

            if (!datasetId.HasValue) throw new Exception("Dataset ID required for Power BI values.");

            var daxQuery = $"EVALUATE DISTINCT('{tableName}'[{columnName}])";
            
            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($">>>> [FLOW] Starting DAX Probe for Power BI report...");
            Console.WriteLine($"[DYNAMIC] TARGET: '{tableName}'[{columnName}]");
            Console.WriteLine($"[DYNAMIC] DAX: {daxQuery}");

            var results = new List<string>();
            try
            {
                var client = await GetClient();
                var request = new DatasetExecuteQueriesRequest(new List<DatasetExecuteQueriesQuery> { new DatasetExecuteQueriesQuery(daxQuery) });
                var response = await client.Datasets.ExecuteQueriesAsync(datasetId.ToString(), request);

                if (response?.Results != null && response.Results.Count > 0 && response.Results[0].Tables != null && response.Results[0].Tables!.Count > 0)
                {
                    var firstResult = response.Results[0];
                    if (firstResult.Tables != null && firstResult.Tables.Count > 0)
                    {
                        var table = firstResult.Tables[0];
                        if (table.Rows != null)
                        {
                            foreach (dynamic row in table.Rows)
                            {
                                var rowStr = row?.ToString();
                                if (string.IsNullOrEmpty(rowStr)) continue;

                                var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(rowStr);
                                if (dict != null && dict.Values.Count > 0)
                                {
                                    foreach (var valObj in dict.Values)
                                    {
                                        string? valStr = valObj?.ToString();
                                        if (!string.IsNullOrEmpty(valStr))
                                        {
                                            results.Add(valStr);
                                            break; 
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                Console.WriteLine($">>>> [FLOW] DAX Result: {results.Count} values found.");
            }
            catch (Exception ex) 
            { 
                if (ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
                {
                    Console.WriteLine("[DYNAMIC] 403 ERROR: Service Principal requires Premium/Fabric capacity to access dataset metadata/queries.");
                }
                Console.WriteLine($">>>> [FLOW] DAX Error: {ex.Message}"); 
            }
            return results;
        }


        public async Task<int> GetOrCreateFolder(int localWorkspaceId, Guid pbiWorkspaceId, string folderName, AppDbContext db)
        {
            Console.WriteLine($"[SERVICE] GetOrCreateFolder: Looking for '{folderName}' in workspace {pbiWorkspaceId}");
            
            // 1. Check local DB first
            var localFolder = await db.Folders.FirstOrDefaultAsync(f => f.WorkspaceId == localWorkspaceId && f.Name == folderName);
            if (localFolder != null) return localFolder.Id;

            var token = await _auth.GetFabricToken();
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            // 2. Check if folder exists using Fabric API
            string? fabricFolderId = null;
            var listUrl = $"https://api.fabric.microsoft.com/v1/workspaces/{pbiWorkspaceId}/folders";
            var listResponse = await client.GetAsync(listUrl);
            if (listResponse.IsSuccessStatusCode)
            {
                var content = await listResponse.Content.ReadAsStringAsync();
                var folders = JsonConvert.DeserializeObject<dynamic>(content);
                if (folders != null && folders.value != null)
                {
                    foreach (var folder in folders.value)
                    {
                        if (folder != null && folder.displayName != null && folder.displayName.ToString() == folderName) 
                        {
                            fabricFolderId = folder.id.ToString();
                            break;
                        }
                    }
                }
            }

            // 3. Create if not found using Fabric API
            if (string.IsNullOrEmpty(fabricFolderId))
            {
                Console.WriteLine($"[SERVICE] Folder '{folderName}' not found in Fabric. Creating...");
                var createUrl = $"https://api.fabric.microsoft.com/v1/workspaces/{pbiWorkspaceId}/folders";
                var body = JsonConvert.SerializeObject(new { displayName = folderName });
                var createResponse = await client.PostAsync(createUrl, new StringContent(body, Encoding.UTF8, "application/json"));
                
                var resultJson = await createResponse.Content.ReadAsStringAsync();
                if (createResponse.IsSuccessStatusCode)
                {
                    var folder = JsonConvert.DeserializeObject<dynamic>(resultJson);
                    fabricFolderId = folder?.id?.ToString();
                }
                else
                {
                     throw new Exception($"Fabric Folder API failed: {resultJson}");
                }
            }

            // 4. Save/Sync to local DB and return ID
            if (!string.IsNullOrEmpty(fabricFolderId))
            {
                var folder = new Folder
                {
                    Name = folderName,
                    FabricFolderId = fabricFolderId,
                    WorkspaceId = localWorkspaceId
                };
                db.Folders.Add(folder);
                await db.SaveChangesAsync();
                return folder.Id;
            }

            return 0;
        }

        private async Task MoveItemToFolder(Guid workspaceId, Guid itemId, Guid targetFolderId)
        {
            var token = await _auth.GetFabricToken();
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var url = $"https://api.fabric.microsoft.com/v1/workspaces/{workspaceId}/items/{itemId}/move";
            var body = JsonConvert.SerializeObject(new { targetFolderId = targetFolderId.ToString() });

            var response = await client.PostAsync(url, new StringContent(body, Encoding.UTF8, "application/json"));
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new Exception($"Fabric Move API failed with status {(int)response.StatusCode}: {error}");
            }
        }
        public async Task<List<dynamic>> GetDatasetTablesAndColumns(Guid datasetId)
        {
            Console.WriteLine($"[SERVICE] Fetching schema for Dataset: {datasetId}");
            var token = await _auth.GetAccessToken();
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var url = $"https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/tables";

            var response = await httpClient.GetAsync(url);
            var rawJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode) 
            {
                Console.WriteLine($"[SERVICE] API FAIL: {response.StatusCode} - {rawJson}");
                if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    throw new Exception("Access Denied (403). This usually means: 1. The Workspace is NOT on a Premium/Fabric capacity. 2. 'Enhanced Metadata Scan' is disabled for Service Principals in the Admin Portal. 3. The SP is not a Member/Admin of the workspace.");
                }
                throw new Exception($"Power BI API returned {response.StatusCode}: {rawJson}");
            }

            var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
            var tables = root["value"] as Newtonsoft.Json.Linq.JArray;

            var result = new List<dynamic>();
            if (tables != null)
            {
                foreach (var tbl in tables)
                {
                    string? tableName = tbl["name"]?.ToString();
                    if (string.IsNullOrEmpty(tableName)) continue;

                    var cols = tbl["columns"] as Newtonsoft.Json.Linq.JArray;
                    if (cols != null)
                    {
                        foreach (var col in cols)
                        {
                            string? colName = col["name"]?.ToString();
                            if (!string.IsNullOrEmpty(colName))
                            {
                                result.Add(new { table = tableName, column = colName });
                            }
                        }
                    }
                }
            }
            Console.WriteLine($"[SERVICE] Discovery complete. Found {result.Count} columns across all tables.");
            return result;
        }
        public async Task DeleteReport(Guid workspaceId, Guid reportId)
        {
            Console.WriteLine($"[SERVICE] Deleting Report {reportId} from Workspace {workspaceId}");
            var client = await GetClient();
            await client.Reports.DeleteReportInGroupAsync(workspaceId, reportId);
        }

        public async Task DeleteFabricItem(Guid workspaceId, Guid itemId)
        {
            Console.WriteLine($"[SERVICE] Deleting Fabric Item {itemId} from Workspace {workspaceId}");
            var token = await _auth.GetAccessToken();
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            
            // Use the Folders API specifically for folders, otherwise Items API
            var url = $"https://api.fabric.microsoft.com/v1/workspaces/{workspaceId}/folders/{itemId}";
            Console.WriteLine($"[SERVICE] Fabric FOLDER DELETE Request: {url}");
            var response = await httpClient.DeleteAsync(url);

            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                // Fallback to Items API if it wasn't a folder or if it's a legacy item
                url = $"https://api.fabric.microsoft.com/v1/workspaces/{workspaceId}/items/{itemId}";
                Console.WriteLine($"[SERVICE] Falling back to Fabric ITEM DELETE Request: {url}");
                response = await httpClient.DeleteAsync(url);
            }
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"[SERVICE] Fabric Delete Failed ({response.StatusCode}): {error}");
                throw new Exception($"Fabric API error: {response.StatusCode} - {error}");
            }
            Console.WriteLine("[SERVICE] Fabric Item deleted successfully.");
        }

        public async Task<(bool Success, string Message)> InjectFilterToRdl(int reportId, string fieldName, AppDbContext db)
        {
            try
            {
                var report = await db.Reports.FindAsync(reportId);
                if (report == null || report.ReportType != "RDL") return (false, "Not an RDL report");

                Console.WriteLine($">>>> [RDL-INJECT] Starting filter injection for Report '{report.Name}', Field: '{fieldName}'");
                XDocument xDoc;
                if (!string.IsNullOrEmpty(report.RdlContent))
                {
                    Console.WriteLine(">>>> [RDL-INJECT] Loading RDL source from Database (SQL Storage).");
                    xDoc = XDocument.Parse(report.RdlContent);
                }
                else
                {
                    Console.WriteLine(">>>> [RDL-INJECT] RDL Content missing from DB. Falling back to Uploads folder...");
                    var reportName = report.Name ?? "Unknown";
                    var fileName = reportName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? reportName : $"{reportName}.rdl";
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);
                    if (!File.Exists(filePath)) 
                    {
                        Console.WriteLine($">>>> [RDL-INJECT] ERROR: Source file '{fileName}' missing.");
                        return (false, "RDL source missing.");
                    }
                    xDoc = XDocument.Load(filePath);
                }

                var ns = xDoc.Root.GetDefaultNamespace();

                // --- INJECT CONNECTION STRING (ONLY FOR SQL) ---
                string? defaultConn = _config.GetConnectionString("DefaultConnection");
                if (!string.IsNullOrEmpty(defaultConn))
                {
                    foreach (var ds in xDoc.Descendants(ns + "DataSource"))
                    {
                        var connProps = ds.Element(ns + "ConnectionProperties");
                        var provider = connProps?.Element(ns + "DataProvider");
                        
                        // Only update if it's already SQL or if it's missing (defaulting to SQL)
                        // Do NOT touch ENTERDATA as it breaks static reports.
                        if (provider != null && provider.Value.Contains("SQL", StringComparison.OrdinalIgnoreCase))
                        {
                            var connectString = connProps.Element(ns + "ConnectString");
                            if (connectString == null) { connectString = new XElement(ns + "ConnectString"); connProps.Add(connectString); }
                            if (connectString.Value != defaultConn)
                            {
                                connectString.Value = defaultConn;
                                Console.WriteLine($">>>> [RDL-INJECT] ConnectionString updated for SQL DataSource '{ds.Attribute("Name")?.Value}'.");
                            }
                        }
                    }
                }

                // 0. Global Cleanup - REMOVED: Destructive QueryParameters.Remove()
                // We now handle query parameters per-dataset below.

                var filterDef = await db.ReportFilters.FirstOrDefaultAsync(f => f.ReportId == reportId && f.ColumnName == fieldName);
                bool isStrict = filterDef?.IsStrict ?? false;
                Console.WriteLine($">>>> [RDL-INJECT] Mode: {(isStrict ? "STRICT (Query-Level)" : "STANDARD (Dataset-Level)")}");

                // 1. Parameter Injection & Global Nullability Fix
                var parametersNode = xDoc.Descendants(ns + "ReportParameters").FirstOrDefault();
                if (parametersNode == null) { parametersNode = new XElement(ns + "ReportParameters"); xDoc.Root.Add(parametersNode); }

                // ENSURE ALL PARAMETERS ARE NULLABLE (Prevents blank screens on first load)
                foreach (var p in parametersNode.Descendants(ns + "ReportParameter"))
                {
                    if (p.Element(ns + "Nullable") == null) p.Add(new XElement(ns + "Nullable", "true"));
                    else p.Element(ns + "Nullable").Value = "true";

                    if (p.Element(ns + "AllowBlank") == null) p.Add(new XElement(ns + "AllowBlank", "true"));
                    else p.Element(ns + "AllowBlank").Value = "true";
                }

                var existingParam = parametersNode.Descendants(ns + "ReportParameter")
                    .FirstOrDefault(p => p.Attribute("Name")?.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase) == true);

                string actualParamName = fieldName;
                if (existingParam == null)
                {
                    Console.WriteLine($">>>> [RDL-INJECT] Injecting ReportParameter: '{fieldName}'");
                    parametersNode.Add(new XElement(ns + "ReportParameter", new XAttribute("Name", fieldName),
                        new XElement(ns + "DataType", "String"),
                        new XElement(ns + "Nullable", "true"),
                        new XElement(ns + "AllowBlank", "true"),
                        new XElement(ns + "Prompt", fieldName),
                        new XElement(ns + "Hidden", "false")
                    ));
                }
                else 
                {
                    Console.WriteLine($">>>> [RDL-INJECT] ReportParameter '{fieldName}' already exists.");
                    actualParamName = existingParam.Attribute("Name")?.Value ?? fieldName;
                }

                // 2. Dataset Injection
                var datasets = xDoc.Descendants(ns + "DataSet").ToList();
                foreach (var dataset in datasets)
                {
                    var dsName = dataset.Attribute("Name")?.Value;
                    if (string.IsNullOrEmpty(dsName) || 
                        dsName.EndsWith("List", StringComparison.OrdinalIgnoreCase) ||
                        dsName.EndsWith("Values", StringComparison.OrdinalIgnoreCase)) continue;

                    var targetField = dataset.Element(ns + "Fields")?.Elements(ns + "Field")
                        .FirstOrDefault(f => f.Attribute("Name")?.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase) == true);

                    if (targetField != null)
                    {
                        string actualFieldName = targetField.Attribute("Name")?.Value ?? fieldName;

                        if (isStrict)
                        {
                            var queryNode = dataset.Element(ns + "Query");
                            if (queryNode != null)
                            {
                                var queryParams = queryNode.Element(ns + "QueryParameters");
                                if (queryParams == null) { queryParams = new XElement(ns + "QueryParameters"); queryNode.Add(queryParams); }

                                var existingQP = queryParams.Elements(ns + "QueryParameter")
                                    .FirstOrDefault(qp => qp.Attribute("Name")?.Value == "@" + actualFieldName);
                                if (existingQP == null)
                                {
                                    Console.WriteLine($">>>> [STRICT-SETUP] Adding QueryParameter '@{actualFieldName}' to DataSet '{dsName}'.");
                                    queryParams.Add(new XElement(ns + "QueryParameter", new XAttribute("Name", "@" + actualFieldName),
                                        new XElement(ns + "Value", $"=Parameters!{actualParamName}.Value")));
                                }

                                var cmdText = queryNode.Element(ns + "CommandText");
                                if (cmdText != null && !cmdText.Value.Contains("ENTERDATA"))
                                {
                                    string sql = cmdText.Value;
                                    string clause = sql.Contains("WHERE", StringComparison.OrdinalIgnoreCase) ? $" AND {actualFieldName} = @{actualFieldName}" : $" WHERE {actualFieldName} = @{actualFieldName}";
                                    if (!sql.Contains($"@{actualFieldName}")) 
                                    {
                                        cmdText.Value = sql + clause;
                                        Console.WriteLine($">>>> [STRICT-SETUP] SQL Query modified with WHERE clause for field '{actualFieldName}'.");
                                    }
                                }
                            }
                        }
                        else
                        {
                            // --- DATASET FILTER INJECTION ---
                            var filters = dataset.Element(ns + "Filters");
                            if (filters == null) { filters = new XElement(ns + "Filters"); dataset.Add(filters); }

                            // Remove existing filters for this parameter to avoid duplicates
                            filters.Elements(ns + "Filter")
                                .Where(f => f.Element(ns + "FilterExpression")?.Value.Contains($"Parameters!{actualParamName}.Value") == true)
                                .Remove();

                            // Simplified but powerful expression: 
                            // If Parameter is blank/null/ALL, it matches everything (Field = Field).
                            // Otherwise, it matches the specific value.
                            string fieldRef = $"Fields!{actualFieldName}.Value";
                            string paramRef = $"Parameters!{actualParamName}.Value";
                            
                            Console.WriteLine($">>>> [UNIVERSAL-SYNC] Applying filter to DataSet '{dsName}' for field '{actualFieldName}'");
                            
                            filters.Add(new XElement(ns + "Filter",
                                new XElement(ns + "FilterExpression", $"={fieldRef}"),
                                new XElement(ns + "Operator", "Equal"),
                                new XElement(ns + "FilterValues", new XElement(ns + "FilterValue", 
                                    $"=IIf({paramRef} = \"\" OR {paramRef} IS NOTHING OR {paramRef} = \"ALL\", {fieldRef}, {paramRef})"))));
                        }
                    }
                }

                // --- LAYOUT CLEANUP (CRITICAL for Power BI Sync) ---
                // We must remove the layout grid before saving/uploading so Power BI can auto-regenerate it 
                // to match our new parameter count. Otherwise, upload fails with rsInvalidParameterLayoutCellDefNotEqualsParameterCount.
                var layout = xDoc.Descendants(ns + "ReportParametersLayout").FirstOrDefault();
                if (layout != null) 
                {
                    Console.WriteLine(">>>> [RDL-INJECT] Removing stale Parameter Layout grid to force regeneration.");
                    layout.Remove();
                }

                // --- SAVE RESULTS TO DATABASE ---
                Console.WriteLine(">>>> [RDL-INJECT] Saving updated RDL Content to Database.");
                report.RdlContent = xDoc.ToString();
                db.Reports.Update(report);
                await db.SaveChangesAsync();

                Console.WriteLine(">>>> [RDL-INJECT] Injection process COMPLETED.");
                return (true, "Success");
            }
            catch (Exception ex) { return (false, ex.Message); }
        }

        public async Task PatchReportCredentials(Guid workspaceId, Guid reportId)
        {
            try
            {
                var client = await GetClient();
                Console.WriteLine($">>>> [PATCH] Starting Credential Patching for Report: {reportId}");
                
                var report = await client.Reports.GetReportInGroupAsync(workspaceId, reportId);
                var datasources = await client.Reports.GetDatasourcesInGroupAsync(workspaceId, reportId);
                
                string? datasetId = report.DatasetId;

                // --- RDL DATASET DISCOVERY FALLBACK ---
                if (string.IsNullOrEmpty(datasetId))
                {
                    var allDatasets = await client.Datasets.GetDatasetsInGroupAsync(workspaceId);
                    var match = allDatasets.Value.FirstOrDefault(d => d.Name == report.Name || d.Name == report.Name + ".rdl");
                    datasetId = match?.Id;
                }

                // --- AZURE-IT-SPEC: Smart Discovery ---
                if (datasources.Value == null || !datasources.Value.Any())
                {
                    var localPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", report.Name.EndsWith(".rdl") ? report.Name : report.Name + ".rdl");
                    if (File.Exists(localPath))
                    {
                        var xDoc = XElement.Load(localPath);
                        var ns = xDoc.GetDefaultNamespace();
                        var dsName = xDoc.Descendants(ns + "DataSource").FirstOrDefault()?.Attribute("Name")?.Value;
                        var provider = xDoc.Descendants(ns + "DataProvider").FirstOrDefault()?.Value;
                        
                        if (!string.IsNullOrEmpty(dsName))
                        {
                            if (provider != null && provider.Contains("ENTERDATA"))
                            {
                                Console.WriteLine($">>>> [PATCH] SKIP: DataSource '{dsName}' is Static XML (ENTERDATA).");
                            }
                            else
                            {
                                dynamic syntheticDs = new System.Dynamic.ExpandoObject();
                                syntheticDs.Name = dsName;
                                syntheticDs.DatasourceType = "Sql";
                                await AttemptAutoGatewayBinding(workspaceId, reportId, datasetId, syntheticDs);
                            }
                        }
                    }
                }
                else 
                {
                    foreach (var reportDs in datasources.Value)
                    {
                        await AttemptAutoGatewayBinding(workspaceId, reportId, datasetId, reportDs);
                    }
                }
                Console.WriteLine(">>>> [PATCH] Process COMPLETED.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($">>>> [PATCH] FAILED: {ex.Message}");
            }
        }

        public class GatewayBindingResult
        {
            public bool Success { get; set; }
            public string? GatewayId { get; set; }
            public string? DatasourceId { get; set; }
        }

        private async Task<List<GatewayBindingResult>> FindAllMatchingGatewayDatasources(Guid workspaceId, string server, string db)
        {
            var results = new List<GatewayBindingResult>();
            try
            {
                var client = await GetClient();
                var gateways = await client.Gateways.GetGatewaysAsync();
                foreach (var gateway in gateways.Value ?? new List<Microsoft.PowerBI.Api.Models.Gateway>())
                {
                    var gwDatasources = await client.Gateways.GetDatasourcesAsync(gateway.Id);
                    foreach (var gds in gwDatasources.Value)
                    {
                        if (gds.DatasourceType == "Sql")
                        {
                            string gwConn = (gds.ConnectionDetails ?? "").Replace("\\\\", "\\");
                            if (gwConn.Contains(server, StringComparison.OrdinalIgnoreCase) && 
                                gwConn.Contains(db, StringComparison.OrdinalIgnoreCase))
                            {
                                results.Add(new GatewayBindingResult { 
                                    Success = true, 
                                    GatewayId = gateway.Id.ToString(), 
                                    DatasourceId = gds.Id.ToString() 
                                });
                            }
                        }
                    }
                }
            }
            catch { }
            return results;
        }

        private async Task<GatewayBindingResult> AttemptAutoGatewayBinding(Guid workspaceId, Guid reportId, string? datasetId, dynamic? reportDs)
        {
            try
            {
                // Fallback: If datasetId is null, some APIs allow using reportId for binding RDLs
                string targetId = datasetId ?? reportId.ToString();
                
                var client = await GetClient();
                Console.WriteLine($">>>> [GATEWAY] Attempting binding for Target ID: {targetId}");

                var gateways = await client.Gateways.GetGatewaysAsync();
                if (gateways.Value == null || !gateways.Value.Any())
                {
                    Console.WriteLine(">>>> [GATEWAY] WARNING: The Service Principal has NO access to any Gateways. Please add the App ID to the 'Manage Users' section of the Gateway in Power BI Portal.");
                }
                else
                {
                    Console.WriteLine($">>>> [GATEWAY] Service Principal has access to {gateways.Value.Count} Gateway(s). Scanning for matches...");
                }

                foreach (var gateway in gateways.Value ?? new List<Microsoft.PowerBI.Api.Models.Gateway>())
                {
                    Console.WriteLine($">>>> [GATEWAY] Checking Gateway Cluster: '{gateway.Name}' ({gateway.Id})");
                    var gwDatasources = await client.Gateways.GetDatasourcesAsync(gateway.Id);
                    
                    Console.WriteLine($">>>> [GATEWAY]   Found {gwDatasources.Value?.Count ?? 0} connections on this cluster.");
                    foreach (var gds in gwDatasources.Value)
                    {
                        Console.WriteLine($">>>> [GATEWAY]   - Connection: '{gds.DatasourceName}' (Type: {gds.DatasourceType})");
                        
                        string targetServer = "___";
                        string targetDb = "___";
                        bool match = false;

                        if (gds.DatasourceType == "Sql")
                        {
                            if (reportDs != null)
                            {
                                try { targetServer = (reportDs.ConnectionDetails?.Server ?? "___").Replace("\\\\", "\\"); } catch { }
                                try { targetDb = (reportDs.ConnectionDetails?.Database ?? "___").Replace("\\\\", "\\"); } catch { }
                            }

                            // CRITICAL FALLBACK: If we still have underscores, use our known local server
                            if (targetServer == "___") targetServer = "zsdeskt8ibnl5\\sqlexpress";
                            if (targetDb == "___") targetDb = "powerBIdemo";

                            string gwConn = (gds.ConnectionDetails ?? "").Replace("\\\\", "\\");
                            
                            Console.WriteLine($">>>> [GATEWAY]     Comparing: Target(Server={targetServer}, DB={targetDb}) vs GatewayDetails({gwConn})");
                            
                            match = gwConn.Contains(targetServer, StringComparison.OrdinalIgnoreCase) &&
                                    gwConn.Contains(targetDb, StringComparison.OrdinalIgnoreCase);
                            
                            // Secondary Match: Flexible name match
                            if (!match && gwConn.Contains(targetServer, StringComparison.OrdinalIgnoreCase))
                            {
                                if (gds.DatasourceName.Contains("PowerBI", StringComparison.OrdinalIgnoreCase) || 
                                    gds.DatasourceName.Contains("Demo", StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($">>>> [GATEWAY] ALERT: DB name mismatch, but Server and Connection Name match. Proceeding...");
                                    match = true;
                                }
                            }
                        }

                        if (match)
                        {
                            Console.WriteLine($">>>> [GATEWAY] MATCH CONFIRMED: Gateway='{gateway.Name}', Connection='{gds.DatasourceName}'");
                            
                            // [SYNC] Force exact casing from Gateway to avoid 404/403 data source errors
                            try 
                            {
                                if (!string.IsNullOrEmpty(gds.ConnectionDetails))
                                {
                                    var gwDetails = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(gds.ConnectionDetails);
                                    if (gwDetails.TryGetProperty("server", out var sProp)) targetServer = sProp.GetString() ?? targetServer;
                                    if (gwDetails.TryGetProperty("database", out var dProp)) targetDb = dProp.GetString() ?? targetDb;
                                    Console.WriteLine($">>>> [GATEWAY] SYNC: Using Gateway Casing -> Server: {targetServer}, DB: {targetDb}");
                                }
                            } catch { }
                            
                            // [AZURE-IT-SPEC] - Raw REST API Fallback to bypass SDK version issues
                            try 
                            {
                                Console.WriteLine(">>>> [GATEWAY] Sending Raw REST API request for RDL Binding...");
                                var token = await _auth.GetAccessToken();
                                using (var httpClient = new System.Net.Http.HttpClient())
                                {
                                    httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                                    
                                    // Parse the connection details string into a real object so it's not double-serialized
                                    object? connObj = null;
                                    try { 
                                        if (!string.IsNullOrEmpty(gds.ConnectionDetails))
                                            connObj = System.Text.Json.JsonSerializer.Deserialize<object>(gds.ConnectionDetails); 
                                        else
                                            connObj = new { };
                                    }
                                    catch { connObj = gds.ConnectionDetails ?? (object)new { }; }

                                    var payload = new
                                    {
                                        updateDetails = new[]
                                        {
                                            new
                                            {
                                                datasourceName = (string)(reportDs?.Name ?? "DataSource"),
                                                gatewayId = gateway.Id.ToString(),
                                                datasourceId = gds.Id.ToString(),
                                                connectionDetails = new {
                                                    server = targetServer,
                                                    database = targetDb
                                                }
                                            }
                                        }
                                    };

                                    var json = System.Text.Json.JsonSerializer.Serialize(payload, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                                    Console.WriteLine($">>>> [GATEWAY] Sending Payload:\n{json}");

                                    var content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
                                    
                                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspaceId}/reports/{reportId}/Default.UpdateDatasources";
                                    var response = await httpClient.PostAsync(url, content);
                                    
                                    if (response.IsSuccessStatusCode)
                                    {
                                        Console.WriteLine(">>>> [GATEWAY] SUCCESS: RDL bound via Raw REST API.");
                                        return new GatewayBindingResult { Success = true, GatewayId = gateway.Id.ToString(), DatasourceId = gds.Id.ToString() };
                                    }
                                    else
                                    {
                                        var error = await response.Content.ReadAsStringAsync();
                                        Console.WriteLine($">>>> [GATEWAY] FAILED: {error}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($">>>> [GATEWAY] ERROR: {ex.Message}");
                            }

                            // Dataset Binding Fallback (Standard Power BI)
                            if (!string.IsNullOrEmpty(datasetId))
                            {
                                try {
                                    var bindRequest = new Microsoft.PowerBI.Api.Models.BindToGatewayRequest(gateway.Id, new List<Guid?> { gds.Id });
                                    await client.Datasets.BindToGatewayInGroupAsync(workspaceId, datasetId, bindRequest);
                                    Console.WriteLine($">>>> [GATEWAY] SUCCESS: Bound via Dataset API fallback.");
                                    return new GatewayBindingResult { Success = true, GatewayId = gateway.Id.ToString(), DatasourceId = gds.Id.ToString() };
                                } catch { /* Ignore dataset errors if RDL */ }
                            }
                        }
                    }
                }

                // --- FALLTHROUGH: CLOUD PATCHING ---
                if (reportDs != null && (string)reportDs.DatasourceType == "Sql")
                {
                    Console.WriteLine(">>>> [GATEWAY] No gateway match. Attempting Cloud Credential Patching...");
                }
                
                Console.WriteLine(">>>> [GATEWAY] No matching gateway connection found.");
                return new GatewayBindingResult { Success = false };
            }
            catch (Exception ex)
            {
                Console.WriteLine($">>>> [GATEWAY] ERROR: {ex.Message}");
                return new GatewayBindingResult { Success = false };
            }
        }

    }
}