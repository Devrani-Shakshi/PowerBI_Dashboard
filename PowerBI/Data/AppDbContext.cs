namespace PowerBI.Data
{
    using Microsoft.EntityFrameworkCore;
    using PowerBI.Models;

    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions options) : base(options) { }

        public DbSet<User> Users { get; set; }
        public DbSet<Workspace> Workspaces { get; set; }
        public DbSet<Report> Reports { get; set; }
        public DbSet<Folder> Folders { get; set; }
        public DbSet<ReportFilter> ReportFilters { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
        }
    }
}