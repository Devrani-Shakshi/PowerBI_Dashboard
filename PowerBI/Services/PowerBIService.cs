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

namespace PowerBI.Services
{
    public class PowerBIService
    {
        private readonly PowerBIAuthService _auth;

        public PowerBIService(PowerBIAuthService auth)
        {
            _auth = auth;
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

        // Get or Create Workspace
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

                var localWorkspaces = db.Workspaces.Where(w => w.UserId == userId).ToList();


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

        public async Task SyncReports(int localWorkspaceId, Guid pbiWorkspaceId, AppDbContext db)
        {
            try 
            {
                // 1. Sync Folders First
                await SyncFolders(localWorkspaceId, pbiWorkspaceId, db);

                Console.WriteLine($"[SERVICE] Syncing reports for Workspace: {pbiWorkspaceId}");
                var client = await GetClient();
                var pbiReports = (await client.Reports.GetReportsInGroupAsync(pbiWorkspaceId)).Value;

                // 2. Fetch item metadata from Fabric to get folder associations
                var fabricToken = await _auth.GetFabricToken();
                using var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", fabricToken);
                
                // Explicitly set recursive=true to ensure we see items inside folders
                var itemsUrl = $"https://api.fabric.microsoft.com/v1/workspaces/{pbiWorkspaceId}/items?recursive=true";
                var itemsResponse = await httpClient.GetAsync(itemsUrl);
                var itemsJson = await itemsResponse.Content.ReadAsStringAsync();
                
                if (!itemsResponse.IsSuccessStatusCode)
                {
                    Console.WriteLine($"[SYNC] ERROR: Fabric Items API failed with {itemsResponse.StatusCode}. Content: {itemsJson}");
                }

                var itemsRoot = JsonConvert.DeserializeObject<dynamic>(itemsJson);
                var fabricItems = itemsRoot?.value;

                if (fabricItems != null)
                {
                    var itemList = (Newtonsoft.Json.Linq.JArray)fabricItems;
                    Console.WriteLine($"[SYNC] Fabric API returned {itemList.Count} items.");
                    foreach (var item in fabricItems)
                    {
                        Console.WriteLine($"[SYNC] Found Fabric Item: {item.displayName} (Type: {item.type}, Parent: {item.parentFolderId ?? "Root"})");
                    }
                }
                else if (itemsResponse.IsSuccessStatusCode)
                {
                    Console.WriteLine("[SYNC] WARNING: Fabric API returned success but 'value' was null or empty.");
                }

                var localReports = db.Reports.Where(r => r.WorkspaceId == localWorkspaceId).ToList();
                var localFolders = db.Folders.Where(f => f.WorkspaceId == localWorkspaceId).ToList();

                foreach (var pbiRep in pbiReports)
                {
                    // Find parent folder from Fabric Items
                    string? parentFolderId = null;
                    if (fabricItems != null)
                    {
                        foreach (var item in fabricItems)
                        {
                            if (item.id?.ToString().Equals(pbiRep.Id.ToString(), StringComparison.OrdinalIgnoreCase) == true)
                            {
                                parentFolderId = item.parentFolderId?.ToString();
                                break;
                            }
                        }
                    }

                    int? localFolderId = null;
                    if (!string.IsNullOrEmpty(parentFolderId))
                    {
                        var folderMatch = localFolders.FirstOrDefault(f => f.FabricFolderId == parentFolderId);
                        localFolderId = folderMatch?.Id;
                        Console.WriteLine($"[SYNC] Report '{pbiRep.Name}' matched to Folder: '{folderMatch?.Name ?? "Unknown"}' (Fabric ID: {parentFolderId})");
                    }
                    else 
                    {
                        Console.WriteLine($"[SYNC] Report '{pbiRep.Name}' has NO parent folder (Root).");
                    }

                    var existing = localReports.FirstOrDefault(r => r.PowerBIReportId == pbiRep.Id.ToString());
                    if (existing == null)
                    {
                        Console.WriteLine($"[SERVICE] Adding new PBI report to DB: {pbiRep.Name} (Folder: {parentFolderId})");
                        db.Reports.Add(new PowerBI.Models.Report
                        {
                            Name = pbiRep.Name,
                            PowerBIReportId = pbiRep.Id.ToString(),
                            PowerBIDatasetId = pbiRep.DatasetId,
                            WorkspaceId = localWorkspaceId,
                            FolderId = localFolderId,
                            ReportType = pbiRep.ReportType == "PaginatedReport" ? "RDL" : "PowerBI"
                        });
                    }
                    else
                    {
                        if (existing.Name != pbiRep.Name) existing.Name = pbiRep.Name;
                        var correctType = pbiRep.ReportType == "PaginatedReport" ? "RDL" : "PowerBI";
                        if (existing.ReportType != correctType) existing.ReportType = correctType;
                        // Smart Sync: Only overwrite local FolderId if the API gives us a non-null folder,
                        // OR if we don't already have one locally. This prevents API lag from 
                        // moving items back to the Root accidentally.
                        if (localFolderId != null || existing.FolderId == null)
                        {
                            existing.FolderId = localFolderId;
                        }
                    }
                }

                foreach (var localRep in localReports)
                {
                    if (!pbiReports.Any(pbi => pbi.Id.ToString() == localRep.PowerBIReportId))
                    {
                        Console.WriteLine($"[SERVICE] Removing report from DB (deleted in PBI): {localRep.Name}");
                        db.Reports.Remove(localRep);
                    }
                }

                await db.SaveChangesAsync();
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] Power BI API Error during SyncReports: {ex.Response.Content}");
                throw new Exception($"Power BI API Error: {ex.Response.ReasonPhrase}. Details: {ex.Response.Content}");
            }
        }

        public async Task<Import> UploadReport(int localWorkspaceId, Guid pbiWorkspaceId, string name, Stream stream, AppDbContext db, string? folderName = null)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Importing {name} to Workspace {pbiWorkspaceId}");
                var client = await GetClient();

                if (name.EndsWith(".pbit", StringComparison.OrdinalIgnoreCase))
                {
                    throw new Exception("Power BI Templates (.pbit) are not supported for direct upload. Please save your file as a .pbix and try again.");
                }

                var isRdl = name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase);
                var datasetName = isRdl ? name : Path.GetFileNameWithoutExtension(name);
                var targetFolder = string.IsNullOrEmpty(folderName) ? "Automated Reports" : folderName;

                Console.WriteLine($"[SERVICE] Stream length: {stream.Length} bytes");

                if (stream.Length == 0)
                    throw new Exception("The uploaded file is empty.");

                var conflictMode = isRdl ? ImportConflictHandlerMode.Abort : ImportConflictHandlerMode.CreateOrOverwrite;
                var finalDisplayName = datasetName;

                Console.WriteLine($"[SERVICE] Uploading {name} as {finalDisplayName} (Mode: {conflictMode})");

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

                Import import;
                
                // Step 1: Upload to Root
                Console.WriteLine("[SERVICE] Uploading to workspace root...");
                import = await client.Imports.PostImportWithFileAsync(
                    pbiWorkspaceId,
                    stream,
                    datasetDisplayName: finalDisplayName,
                    nameConflict: conflictMode
                );

                // Step 2: Poll for completion
                Console.WriteLine($"[SERVICE] Import '{import.Id}' started. Waiting for completion...");
                int attempts = 0;
                while (import.ImportState != "Succeeded" && import.ImportState != "Failed" && attempts < 15)
                {
                    await Task.Delay(3000);
                    import = await client.Imports.GetImportInGroupAsync(pbiWorkspaceId, import.Id);
                    Console.WriteLine($"[SERVICE] Import Status: {import.ImportState}...");
                    attempts++;
                }

                if (import.ImportState == "Failed")
                {
                    Console.WriteLine($"[SERVICE] ERROR: Power BI Import failed. State: {import.ImportState}");
                    throw new Exception("Power BI Import failed. This can happen if the file is corrupted or uses unsupported features.");
                }

                // Step 3: Move to Folder if discovered
                if (folderId.HasValue && import.Reports != null && import.Reports.Any())
                {
                    var reportId = import.Reports.First().Id;
                    try 
                    {
                        var folder = await db.Folders.FindAsync(folderId.Value);
                        if (folder != null && !string.IsNullOrEmpty(folder.FabricFolderId))
                        {
                            Console.WriteLine($"[SERVICE] Moving Report {reportId} to Fabric Folder {folder.FabricFolderId}...");
                            await MoveItemToFolder(pbiWorkspaceId, reportId, Guid.Parse(folder.FabricFolderId));
                            Console.WriteLine("[SERVICE] SUCCESS: Report moved to folder.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SERVICE] Warning: Could not move report to folder: {ex.Message}");
                    }
                }

                return import;
                // --- FOLDER LOGIC END ---
            }
            catch (Exception ex)
            {
                throw new Exception($"Upload failed: {ex.Message}");
            }
        }

        public async Task<Stream> ExportReportAsStream(Guid workspaceId, Guid reportId, List<ExportFilter>? filters = null)
        {
            Console.WriteLine($"[SERVICE] Exporting Report {reportId} to PDF (Filters: {filters?.Count ?? 0})...");
            var client = await GetClient();

            var exportRequest = new ExportReportRequest 
            { 
                Format = FileFormat.PDF,
                PowerBIReportConfiguration = new Microsoft.PowerBI.Api.Models.PowerBIReportExportConfiguration
                {
                    ReportLevelFilters = filters
                }
            };
            
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

        public async Task<ReportEmbedConfig> GetEmbedConfig(Guid workspaceId, Guid reportId)
        {
            Console.WriteLine($"[SERVICE] Fetching Embed Config for Report {reportId}");
            var client = await GetClient();

            var report = await client.Reports.GetReportInGroupAsync(workspaceId, reportId);
            
            var generateTokenRequestParameters = new GenerateTokenRequest(accessLevel: "view");
            var tokenResponse = await client.Reports.GenerateTokenInGroupAsync(workspaceId, reportId, generateTokenRequestParameters);

            Console.WriteLine("[SERVICE] Embed Token generated successfully.");

            return new ReportEmbedConfig
            {
                ReportId = report.Id.ToString(),
                DatasetId = report.DatasetId,
                EmbedUrl = report.EmbedUrl,
                EmbedToken = tokenResponse.Token,
                ReportName = report.Name
            };
        }


        public async Task DiscoverReportFilters(int reportId, Guid? datasetId, AppDbContext db)
        {
            var report = await db.Reports.FindAsync(reportId);
            if (report == null) throw new Exception("Report not found in local DB.");

            var workspace = await db.Workspaces.FindAsync(report.WorkspaceId);
            if (workspace == null || string.IsNullOrEmpty(workspace.PowerBIWorkspaceId))
                throw new Exception("Workspace not found or not synced.");

            Console.WriteLine($"[SCHEMA-DISCOVERY] MODE: {report.ReportType}");
            Console.WriteLine($"[SCHEMA-DISCOVERY] Report Name: {report.Name}");

            var existing = db.ReportFilters.Where(f => f.ReportId == reportId).ToList();
            Console.WriteLine($"[SCHEMA-DISCOVERY] Purging {existing.Count} stale records for reportId={reportId}");
            db.ReportFilters.RemoveRange(existing);
            await db.SaveChangesAsync();

            int count = 0;

            if (report.ReportType == "RDL")
            {
                try
                {
                    Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: Attempting API discovery...");
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspace.PowerBIWorkspaceId}/reports/{report.PowerBIReportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    if (response.IsSuccessStatusCode)
                    {
                        var rawJson = await response.Content.ReadAsStringAsync();
                        var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                        var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                string? paramName = param["name"]?.ToString();
                                if (string.IsNullOrEmpty(paramName)) continue;

                                db.ReportFilters.Add(new ReportFilter
                                {
                                    ReportId = reportId,
                                    TableName = "RDL_PARAMETER",
                                    ColumnName = paramName,
                                    DisplayName = paramName,
                                    IsActive = true
                                });
                                count++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: API returned {response.StatusCode}. Falling back to Local XML Parsing...");
                        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", $"{report.Name}.rdl");
                        if (File.Exists(filePath))
                        {
                            var xml = System.Xml.Linq.XDocument.Load(filePath);
                            var ns = xml.Root?.GetDefaultNamespace();
                            if (ns != null)
                            {
                                var rdlParams = xml.Descendants(ns + "ReportParameter");

                                foreach (var p in rdlParams)
                                {
                                    string? paramName = p.Attribute("Name")?.Value;
                                    if (string.IsNullOrEmpty(paramName)) continue;

                                    db.ReportFilters.Add(new PowerBI.Models.ReportFilter
                                    {
                                        ReportId = reportId,
                                        TableName = "RDL_PARAMETER",
                                        ColumnName = paramName,
                                        DisplayName = paramName,
                                        IsActive = true
                                    });
                                    count++;
                                }
                            }
                            Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: Local parsing found {count} parameters.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SCHEMA-DISCOVERY] RDL Error: {ex.Message}");
                    throw new Exception($"Failed to discover RDL parameters: {ex.Message}");
                }
            }
            else
            {
                if (!datasetId.HasValue) throw new Exception("Dataset ID is missing for Power BI report.");

                string rawJson;
                try
                {
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/tables";

                    Console.WriteLine($"[SCHEMA-DISCOVERY] Calling: GET {url}");
                    var response = await httpClient.GetAsync(url);
                    rawJson = await response.Content.ReadAsStringAsync();

                    if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        throw new Exception("Access Denied (403). For Service Principals, this usually means the workspace is NOT on a Premium/Fabric capacity, or the 'XMLA Endpoint' setting is not 'Read Write'. Please wait a few minutes after assigning the capacity and try again.");
                    }

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"Power BI REST API returned {(int)response.StatusCode}: {rawJson}");

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

                                    db.ReportFilters.Add(new ReportFilter
                                    {
                                        ReportId = reportId,
                                        TableName = tableName,
                                        ColumnName = colName,
                                        DisplayName = colName,
                                        IsActive = count < 6
                                    });
                                    count++;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to discover Power BI schema: {ex.Message}");
                }
            }

            if (count == 0) throw new Exception("No queryable fields or parameters discovered.");

            await db.SaveChangesAsync();
            Console.WriteLine($"[SCHEMA-DISCOVERY] SUCCESS: {count} field(s) persisted to DB.");
        }




        public async Task<List<string>> GetColumnValues(Guid? datasetId, string tableName, string columnName, Guid? reportId = null, Guid? workspaceId = null)
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
                throw new Exception("[DYNAMIC] Error: Missing Table or Column name.");

            if (tableName == "RDL_PARAMETER")
            {
                if (!reportId.HasValue || !workspaceId.HasValue) throw new Exception("Report/Workspace ID required for RDL values.");
                
                try
                {
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspaceId}/reports/{reportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    var rawJson = await response.Content.ReadAsStringAsync();
                    var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                    var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                    if (parameters != null)
                    {
                        var param = parameters.FirstOrDefault(p => p["name"]?.ToString() == columnName);
                        if (param != null)
                        {
                            var suggested = param["suggestedValues"] as Newtonsoft.Json.Linq.JArray;
                            if (suggested != null)
                            {
                                return suggested.Select(v => v.ToString()).ToList();
                            }
                        }
                    }
                    return new List<string>();
                }
                catch { return new List<string>(); }
            }
            if (!datasetId.HasValue) throw new Exception("Dataset ID required for Power BI values.");

            var daxQuery = $"EVALUATE DISTINCT('{tableName}'[{columnName}])";
            
            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"[DYNAMIC] PROBE: '{tableName}'[{columnName}]");
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
                Console.WriteLine($"[DYNAMIC] RESULT: {results.Count} values found.");
            }
            catch (Exception ex) 
            { 
                if (ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
                {
                    Console.WriteLine("[DYNAMIC] 403 ERROR: Service Principal requires Premium/Fabric capacity to access dataset metadata/queries.");
                }
                Console.WriteLine($"[DYNAMIC] DAX FAIL: {ex.Message}"); 
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
                        if (folder?.displayName?.ToString() == folderName) 
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
    }
}