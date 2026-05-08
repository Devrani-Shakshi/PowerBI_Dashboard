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
using System.IO;

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

                var localReports = db.Reports.Where(r => r.WorkspaceId == localWorkspaceId).ToList();
                var localFolders = db.Folders.Where(f => f.WorkspaceId == localWorkspaceId).ToList();

                // 3. Process every item found by Fabric
                var syncedReportIds = new List<string>();
                if (fabricItems != null)
                {
                    foreach (var item in fabricItems)
                    {
                        string type = item.type?.ToString() ?? "";
                        if (type != "Report" && type != "PaginatedReport") continue;

                        string pbiId = item.id.ToString();
                        string displayName = item.displayName.ToString();
                        string? parentFolderId = item.parentFolderId?.ToString();
                        
                        syncedReportIds.Add(pbiId);
                        
                        // Try to find datasetId from pbiReports list
                        var pbiMatch = pbiReports.FirstOrDefault(p => p.Id.ToString().Equals(pbiId, StringComparison.OrdinalIgnoreCase));
                        string? datasetId = pbiMatch?.DatasetId;

                        int? localFolderId = null;
                        if (!string.IsNullOrEmpty(parentFolderId))
                        {
                            var folderMatch = localFolders.FirstOrDefault(f => f.FabricFolderId.Equals(parentFolderId, StringComparison.OrdinalIgnoreCase));
                            localFolderId = folderMatch?.Id;
                            if (localFolderId.HasValue)
                                Console.WriteLine($"[SYNC] Report '{displayName}' belongs to Folder '{folderMatch?.Name}'");
                        }

                        var existing = localReports.FirstOrDefault(r => r.PowerBIReportId.Equals(pbiId, StringComparison.OrdinalIgnoreCase));
                        if (existing == null)
                        {
                            Console.WriteLine($"[SERVICE] Adding new {type} to DB: {displayName} (Folder: {localFolderId})");
                            db.Reports.Add(new PowerBI.Models.Report
                            {
                                Name = displayName,
                                PowerBIReportId = pbiId,
                                PowerBIDatasetId = datasetId,
                                WorkspaceId = localWorkspaceId,
                                FolderId = localFolderId,
                                ReportType = type == "PaginatedReport" ? "RDL" : "PowerBI"
                            });
                        }
                        else
                        {
                            if (existing.Name != displayName) existing.Name = displayName;
                            
                            // Only update folder if we found one in Fabric (prevents "un-moving" if API is slow)
                            if (localFolderId.HasValue) 
                            {
                                existing.FolderId = localFolderId;
                            }
                            
                            if (datasetId != null) existing.PowerBIDatasetId = datasetId;
                            existing.ReportType = type == "PaginatedReport" ? "RDL" : "PowerBI";
                        }
                    }
                }

                // 4. Cleanup: Remove local reports that no longer exist in Fabric
                foreach (var localRep in localReports)
                {
                    if (!syncedReportIds.Contains(localRep.PowerBIReportId))
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
                Console.WriteLine($"[SERVICE] REQUEST: Upload '{name}' (Size: {stream.Length} bytes)");
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

                // RDL usually prefers 'Abort' or 'Ignore', PBIX prefers 'CreateOrOverwrite'
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

                Import import;
                Stream uploadStream = stream;

                // --- RDL NORMALIZATION & INJECTION ---
                if (isRdl)
                {
                    uploadStream = NormalizeAndInjectRdl(stream, name);
                    
                    // CRITICAL: Save the normalized version locally so Discovery/Sidebar can see the injected parameters
                    var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                    if (!Directory.Exists(uploadsDir)) Directory.CreateDirectory(uploadsDir);
                    var filePath = Path.Combine(uploadsDir, name);
                    
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        uploadStream.Position = 0;
                        await uploadStream.CopyToAsync(fileStream);
                        uploadStream.Position = 0;
                    }
                    Console.WriteLine($"[SERVICE] Normalized RDL saved locally: {filePath}");
                }
                
                // Step 1: Upload to Root
                Console.WriteLine($"[SERVICE] Sending file to Power BI: {finalDisplayName} (Size: {uploadStream.Length} bytes)");
                
                try 
                {
                    if (isRdl)
                    {
                        import = await client.Imports.PostImportWithFileAsync(
                            pbiWorkspaceId,
                            uploadStream,
                            datasetDisplayName: finalDisplayName, 
                            nameConflict: conflictMode
                        );
                    }
                    else
                    {
                        import = await client.Imports.PostImportWithFileAsync(
                            pbiWorkspaceId,
                            uploadStream,
                            datasetDisplayName: finalDisplayName,
                            nameConflict: conflictMode
                        );
                    }
                    Console.WriteLine($"[SERVICE] Upload request accepted. Import ID: {import?.Id}");
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
                        }
                        else
                        {
                            Console.WriteLine("[SERVICE] ERROR: Could not find report in workspace after multiple attempts. It may be stuck in Root.");
                        }
                    }
                }

                return import;
            }
            catch (Exception ex)
            {
                throw new Exception($"Upload failed: {ex.Message}");
            }
        }

        // --- RDL XML ENGINE START ---

        private MemoryStream NormalizeAndInjectRdl(Stream inputMetadata, string? rdlName)
        {
            Console.WriteLine($"[RDL-ENGINE] Normalizing XML for {rdlName}...");
            XmlDocument xmlDoc = new XmlDocument();
            
            // Read stream to memory first to avoid closed stream issues
            var mStream = new MemoryStream();
            inputMetadata.CopyTo(mStream);
            mStream.Position = 0;
            xmlDoc.Load(mStream);

            XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
            string nsUri = xmlDoc.DocumentElement.NamespaceURI;
            nsMgr.AddNamespace("rdl", nsUri);

            // 1. DATA SOURCE NORMALIZATION
            XmlNode dataSources = xmlDoc.SelectSingleNode("//rdl:DataSources", nsMgr);
            if (dataSources == null)
            {
                dataSources = docCreateElement(xmlDoc, "DataSources", nsUri);
                xmlDoc.DocumentElement.PrependChild(dataSources);
            }

            foreach (XmlNode ds in dataSources.SelectNodes("rdl:DataSource", nsMgr))
            {
                // Remove shared reference
                XmlNode reference = ds.SelectSingleNode("rdl:DataSourceReference", nsMgr);
                if (reference != null) ds.RemoveChild(reference);

                // Ensure connection properties
                XmlNode connProps = ds.SelectSingleNode("rdl:ConnectionProperties", nsMgr);
                if (connProps == null)
                {
                    connProps = docCreateElement(xmlDoc, "ConnectionProperties", nsUri);
                    ds.AppendChild(connProps);
                }

                // Only force SQLAZURE if it's already a SQL provider
                XmlNode provider = connProps.SelectSingleNode("rdl:DataProvider", nsMgr);
                
                if (provider?.InnerText == "ENTERDATA")
                {
                    Console.WriteLine($"[RDL-ENGINE] Mock Data detected (ENTERDATA). Using DIRECT PASS-THROUGH.");
                    mStream.Position = 0;
                    return mStream;
                }
                
                if (provider?.InnerText == "SQL" || provider?.InnerText == "SQLAZURE")
                {
                    Console.WriteLine($"[RDL-ENGINE] SQL Provider detected: {provider.InnerText} -> Forcing SQLAZURE");
                    provider.InnerText = "SQLAZURE";

                    XmlNode connString = connProps.SelectSingleNode("rdl:ConnectString", nsMgr);
                    if (connString != null && !string.IsNullOrEmpty(connString.InnerText))
                    {
                        Console.WriteLine($"[RDL-ENGINE] Preserving existing connection string: {connString.InnerText}");
                    }
                    else
                    {
                       if (connString == null) 
                       {
                            connString = docCreateElement(xmlDoc, "ConnectString", nsUri);
                            connProps.AppendChild(connString);
                       }
                       Console.WriteLine("[RDL-ENGINE] Connection string empty. Using production fallback.");
                       connString.InnerText = "Data Source=powerbi-prod-server.database.windows.net;Initial Catalog=ReportingDB;";
                    }
                    
                    // 2. DYNAMIC PARAMETER INJECTION (TenantId) - ONLY FOR SQL
                    InjectTenantIdParameter(xmlDoc, nsMgr, nsUri);
                    
                    // 3. ROW LEVEL SECURITY (Inject filter into SQL only)
                    InjectSecurityParameter(xmlDoc, nsMgr, nsUri, "TenantId");
                }
                else
                {
                     Console.WriteLine($"[RDL-ENGINE] Non-SQL Provider detected ({provider?.InnerText ?? "Unknown"}). Using DIRECT PASS-THROUGH.");
                     mStream.Position = 0;
                     return mStream;
                }
            }

            // Step 5: Save to stream WITHOUT BOM (Power BI is picky)
            var outputStream = new MemoryStream();
            var settings = new XmlWriterSettings 
            { 
                Encoding = new UTF8Encoding(false), // false = NO BOM
                Indent = true 
            };
            
            using (var writer = XmlWriter.Create(outputStream, settings))
            {
                xmlDoc.Save(writer);
            }
            
            outputStream.Position = 0;
            return outputStream;
        }

        private XmlElement docCreateElement(XmlDocument doc, string name, string ns)
        {
            return doc.CreateElement(name, ns);
        }

        private void InjectTenantIdParameter(XmlDocument xmlDoc, XmlNamespaceManager nsMgr, string nsUri)
        {
            var parametersNode = xmlDoc.SelectSingleNode("//rdl:ReportParameters", nsMgr);
            if (parametersNode == null)
            {
                parametersNode = docCreateElement(xmlDoc, "ReportParameters", nsUri);
                xmlDoc.DocumentElement.AppendChild(parametersNode);
            }

            if (xmlDoc.SelectSingleNode("//rdl:ReportParameter[@Name='TenantId']", nsMgr) == null)
            {
                Console.WriteLine("[RDL-ENGINE] Injecting security parameter: TenantId");
                var paramNode = docCreateElement(xmlDoc, "ReportParameter", nsUri);
                var nameAttr = xmlDoc.CreateAttribute("Name");
                nameAttr.Value = "TenantId";
                paramNode.Attributes.Append(nameAttr);

                var dataType = docCreateElement(xmlDoc, "DataType", nsUri);
                dataType.InnerText = "String";
                paramNode.AppendChild(dataType);

                var nullable = docCreateElement(xmlDoc, "Nullable", nsUri);
                nullable.InnerText = "true";
                paramNode.AppendChild(nullable);

                var prompt = docCreateElement(xmlDoc, "Prompt", nsUri);
                prompt.InnerText = "TenantId";
                paramNode.AppendChild(prompt);

                parametersNode.AppendChild(paramNode);
            }
        }

        private void InjectSecurityParameter(XmlDocument doc, XmlNamespaceManager nsMgr, string nsUri, string paramName)
        {
            // Check if <ReportParameters> exists
            XmlNode paramsNode = doc.SelectSingleNode("//rdl:ReportParameters", nsMgr);
            if (paramsNode == null)
            {
                paramsNode = docCreateElement(doc, "ReportParameters", nsUri);
                doc.DocumentElement.AppendChild(paramsNode);
            }

            if (paramsNode.SelectSingleNode($"rdl:ReportParameter[@Name='{paramName}']", nsMgr) == null)
            {
                Console.WriteLine($"[RDL-ENGINE] Injecting security parameter: {paramName}");
                XmlElement newParam = docCreateElement(doc, "ReportParameter", nsUri);
                newParam.SetAttribute("Name", paramName);
                newParam.InnerXml = $"<DataType>String</DataType><Nullable>true</Nullable><Prompt>{paramName}</Prompt>";
                paramsNode.AppendChild(newParam);
            }

            // INJECT INTO SQL QUERY (CommandText)
            XmlNodeList queryNodes = doc.SelectNodes("//rdl:Query/rdl:CommandText", nsMgr);
            foreach (XmlNode query in queryNodes)
            {
                string originalSql = query.InnerText;
                if (!originalSql.Contains($"@{paramName}"))
                {
                    Console.WriteLine("[RDL-ENGINE] Appending security filter to SQL CommandText.");
                    if (originalSql.Contains("WHERE", StringComparison.OrdinalIgnoreCase))
                    {
                        query.InnerText = originalSql.Replace("WHERE", $"WHERE {paramName} = @{paramName} AND ", StringComparison.OrdinalIgnoreCase);
                    }
                    else if (originalSql.Contains("ORDER BY", StringComparison.OrdinalIgnoreCase))
                    {
                        query.InnerText = originalSql.Replace("ORDER BY", $"WHERE {paramName} = @{paramName} ORDER BY ", StringComparison.OrdinalIgnoreCase);
                    }
                    else
                    {
                        query.InnerText = originalSql + $" WHERE {paramName} = @{paramName}";
                    }
                }
            }
        }

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
                    if (localRep != null)
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
                        Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: API returned {response.StatusCode}. Processing via XML Parsing...");
                        var fileName = report.Name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? report.Name : $"{report.Name}.rdl";
                        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);
                        
                        if (File.Exists(filePath))
                        {
                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.Load(filePath);
                            XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
                            nsMgr.AddNamespace("rdl", xmlDoc.DocumentElement.NamespaceURI);

                            var rdlParamsList = xmlDoc.SelectNodes("//rdl:ReportParameter", nsMgr);

                            foreach (XmlNode p in rdlParamsList)
                            {
                                string? paramName = p.Attributes?["Name"]?.Value;
                                if (string.IsNullOrEmpty(paramName) || paramName == "TenantId") continue;

                                Console.WriteLine($"[SCHEMA-DISCOVERY] RDL XML: Found Parameter '{paramName}'");
                                db.ReportFilters.Add(new PowerBI.Models.ReportFilter
                                {
                                    ReportId = reportId,
                                    TableName = "RDL_PARAMETER",
                                    ColumnName = paramName, // Preserving case
                                    DisplayName = paramName,
                                    IsActive = true
                                });
                                count++;
                            }
                            Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: XML parsing found {count} parameters.");
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




        public async Task<List<string>> GetColumnValues(Guid? datasetId, string tableName, string columnName, Guid? reportId = null, Guid? workspaceId = null, AppDbContext? db = null)
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
                        // Try alternative endpoint (without /groups/)
                        Console.WriteLine(">>>> [TERMINAL-LOG] Group API failed, trying direct report API...");
                        url = $"https://api.powerbi.com/v1.0/myorg/reports/{reportId}/parameters";
                        response = await httpClient.GetAsync(url);
                    }

                    var rawJson = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($">>>> [TERMINAL-LOG] DEBUG RDL API Response: {rawJson}");
                    
                    if (response.IsSuccessStatusCode)
                    {
                        var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                        var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                if (param["name"]?.ToString() == columnName)
                                {
                                    var suggestedValues = param["suggestedValues"] as Newtonsoft.Json.Linq.JArray;
                                    if (suggestedValues != null)
                                    {
                                        var result = suggestedValues.Select(v => v.ToString()).ToList();
                                        Console.WriteLine($">>>> [TERMINAL-LOG] SUCCESS: Found {result.Count} suggested values for '{columnName}' via API");
                                        return result;
                                    }
                                }
                            }
                        }
                    }

                    // FALLBACK: Parse XML for ValidValues
                    Console.WriteLine($">>>> [TERMINAL-LOG] RDL: API failed or empty. Parsing XML for '{columnName}' values...");
                    
                    var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                    // Try to find the specific file for this report first
                    string? specificFile = null;
                    if (reportId.HasValue && db != null)
                    {
                        var reportRecord = await db.Reports.FirstOrDefaultAsync(r => r.PowerBIReportId == reportId.Value.ToString());
                        if (reportRecord != null)
                        {
                            var fileName = reportRecord.Name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? reportRecord.Name : $"{reportRecord.Name}.rdl";
                            specificFile = Path.Combine(uploadsDir, fileName);
                        }
                    }

                    var filesToScan = specificFile != null && File.Exists(specificFile) 
                        ? new[] { specificFile } 
                        : Directory.GetFiles(uploadsDir, "*.rdl");
                    
                    foreach (var file in filesToScan)
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(file);
                        XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
                        nsMgr.AddNamespace("rdl", xmlDoc.DocumentElement.NamespaceURI);

                        // Find the parameter by name
                        var paramNode = xmlDoc.SelectSingleNode($"//rdl:ReportParameter[@Name='{columnName}']", nsMgr);
                        if (paramNode != null)
                        {
                            var dsRef = paramNode.SelectSingleNode("rdl:ValidValues/rdl:DataSetReference", nsMgr);
                            if (dsRef != null)
                            {
                                Console.WriteLine($"[SERVICE] RDL: Parameter '{columnName}' is Query-based (DataSet: {dsRef.SelectSingleNode("rdl:DataSetName", nsMgr)?.InnerText}). User must type value manually.");
                            }

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
                                Console.WriteLine($">>>> [TERMINAL-LOG] SUCCESS: Found {xmlResults.Count} values in XML for '{columnName}'");
                                return xmlResults.Where(s => !string.IsNullOrEmpty(s)).Distinct().ToList();
                            }
                        }
                    }

                    Console.WriteLine($"[SERVICE] RDL: No suggested values found in API or XML for '{columnName}'.");
                    return new List<string>();
                }
                catch (Exception ex)
                { 
                    Console.WriteLine($">>>> [TERMINAL-LOG] RDL Value Error: {ex.Message}");
                    return new List<string>(); 
                }
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