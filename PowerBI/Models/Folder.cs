namespace PowerBI.Models
{
    public class Folder
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? FabricFolderId { get; set; }
        public int WorkspaceId { get; set; }
    }
}
