namespace PowerBI.Models
{
    public class Report
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? PowerBIReportId { get; set; }
        public string? PowerBIDatasetId { get; set; }
        public int WorkspaceId { get; set; }
        public int? FolderId { get; set; }
        public string? FilePath { get; set; }
        public string? ReportType { get; set; } // "PowerBI" or "RDL"
    }
}
