
using Microsoft.EntityFrameworkCore;
using PowerBI.Models;
using PowerBI.Data;
using PowerBI.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddScoped<PowerBIAuthService>();
builder.Services.AddScoped<PowerBIService>();
builder.Services.AddSession();






var app = builder.Build();

// 🚀 Seed Database
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    // 🛠️ MANUAL SCHEMA PATCH (For Folders feature)
    try 
    {
        Console.WriteLine("[DB] Checking for Folders table...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Folders]') AND type in (N'U'))
            BEGIN
                CREATE TABLE [dbo].[Folders] (
                    [Id] int IDENTITY(1,1) NOT NULL,
                    [Name] nvarchar(max) NULL,
                    [FabricFolderId] nvarchar(max) NULL,
                    [WorkspaceId] int NOT NULL,
                    CONSTRAINT [PK_Folders] PRIMARY KEY CLUSTERED ([Id] ASC)
                );
                PRINT 'Created Folders table.';
            END
        ");

        Console.WriteLine("[DB] Checking for Report.FolderId column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'FolderId' AND Object_ID = OBJECT_ID(N'[dbo].[Reports]'))
            BEGIN
                ALTER TABLE [dbo].[Reports] ADD [FolderId] int NULL;
                PRINT 'Added FolderId column to Reports table.';
            END
        ");

        Console.WriteLine("[DB] Checking for ReportFilters.IsCustom column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'IsCustom' AND Object_ID = OBJECT_ID(N'[dbo].[ReportFilters]'))
            BEGIN
                ALTER TABLE [dbo].[ReportFilters] ADD [IsCustom] bit NOT NULL DEFAULT 0;
                PRINT 'Added IsCustom column to ReportFilters table.';
            END
        ");

        Console.WriteLine("[DB] Checking for Report.CreatedByUserId column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'CreatedByUserId' AND Object_ID = OBJECT_ID(N'[dbo].[Reports]'))
            BEGIN
                ALTER TABLE [dbo].[Reports] ADD [CreatedByUserId] int NULL;
                PRINT 'Added CreatedByUserId column to Reports table.';
            END
        ");

        Console.WriteLine("[DB] Checking for ReportFilters.UserId column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'UserId' AND Object_ID = OBJECT_ID(N'[dbo].[ReportFilters]'))
            BEGIN
                ALTER TABLE [dbo].[ReportFilters] ADD [UserId] int NULL;
                PRINT 'Added UserId column to ReportFilters table.';
            END
        ");

        Console.WriteLine("[DB] Checking for ReportFilters.IsStrict column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'IsStrict' AND Object_ID = OBJECT_ID(N'[dbo].[ReportFilters]'))
            BEGIN
                ALTER TABLE [dbo].[ReportFilters] ADD [IsStrict] bit NOT NULL DEFAULT 0;
                PRINT 'Added IsStrict column to ReportFilters table.';
            END
        ");

        Console.WriteLine("[DB] Checking for Report.RdlContent column...");
        db.Database.ExecuteSqlRaw(@"
            IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'RdlContent' AND Object_ID = OBJECT_ID(N'[dbo].[Reports]'))
            BEGIN
                ALTER TABLE [dbo].[Reports] ADD [RdlContent] nvarchar(max) NULL;
                PRINT 'Added RdlContent column to Reports table.';
            END
        ");

        // UserReportSessions patch removed - Single Report Architecture
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[DB] Patch Warning: {ex.Message}");
    }

    if (!db.Users.Any())
    {
        db.Users.Add(new User 
        { 
            Name = "Admin User", 
            Email = "admin@zealousys.co", 
            Password = "admin", 
            Role = "Admin" 
        });
        db.SaveChanges();
    }

    // 🚀 HOTFIX: Auto-sync all existing custom filters to RDL files to ensure they get the new Bulletproof IIF logic
    try
    {
        Console.WriteLine("[HOTFIX] Checking for RDL files that need to be synced with the new filter engine...");
        var pbiService = scope.ServiceProvider.GetRequiredService<PowerBIService>();
        var rdlReports = db.Reports.Where(r => r.ReportType == "RDL").ToList();
        
        foreach (var report in rdlReports)
        {
            var customFilters = db.ReportFilters.Where(f => f.ReportId == report.Id && f.IsActive && f.ColumnName != "TenantId").ToList();
            bool synced = false;
            foreach (var f in customFilters)
            {
                var result = await pbiService.InjectFilterToRdl(report.Id, f.ColumnName, db);
                if (result.Success) synced = true;
            }

            if (synced)
            {
                var ws = db.Workspaces.Find(report.WorkspaceId);
                if (ws != null && !string.IsNullOrEmpty(ws.PowerBIWorkspaceId))
                {
                    var reportName = report.Name ?? "Unknown";
                    var fileName = reportName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? reportName : $"{reportName}.rdl";
                    var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "wwwroot", "uploads", fileName);
                    
                    if (System.IO.File.Exists(filePath))
                    {
                        byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                        using (var ms = new System.IO.MemoryStream(fileBytes))
                        {
                            Console.WriteLine($"[HOTFIX] Re-uploading {fileName} to Power BI to apply new filter architecture...");
                            var uploadResult = await pbiService.UploadReport(report.WorkspaceId, Guid.Parse(ws.PowerBIWorkspaceId), fileName, ms, db);
                            var importResult = uploadResult.Import;
                            
                            // CRITICAL: Update the database with the NEW Power BI Report ID from the import
                            if (importResult.Reports != null && importResult.Reports.Any())
                            {
                                var newId = importResult.Reports.First().Id.ToString();
                                report.PowerBIReportId = newId;
                                report.RdlContent = uploadResult.RdlContent; // Sync XML to DB as well
                                db.Reports.Update(report);
                                await db.SaveChangesAsync();
                                Console.WriteLine($"[HOTFIX] Database ID updated for {report.Name} to {newId}");
                            }
                        }
                    }
                }
            }
        }
        Console.WriteLine("[HOTFIX] RDL Sync complete.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[HOTFIX] Failed to auto-sync RDLs: {ex.Message}");
    }
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}


app.UseSession();
app.UseHttpsRedirection();
app.UseRouting();

app.UseAuthorization();

app.MapStaticAssets();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Auth}/{action=Login}/{id?}")
    .WithStaticAssets();


app.Run();
