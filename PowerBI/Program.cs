
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
