namespace PowerBI.Models
{
    public class ReportEmbedConfig
    {
        public string? ReportId { get; set; }
        public string? DatasetId { get; set; }
        public string? EmbedUrl { get; set; }
        public string? EmbedToken { get; set; }
        public string? ReportName { get; set; }
        public string? ReportType { get; set; } 
        public int LocalReportId { get; set; }
        public string? WorkspaceId { get; set; }
        public System.Collections.Generic.Dictionary<string, string> Parameters { get; set; } = new();
    }
}
