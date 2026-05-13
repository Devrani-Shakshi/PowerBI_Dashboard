using System.ComponentModel.DataAnnotations;

namespace PowerBI.Models
{
    public class ReportFilter
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int ReportId { get; set; }
        
        [Required]
        public string TableName { get; set; } = string.Empty;
        
        [Required]
        public string ColumnName { get; set; } = string.Empty;
        
        public string DisplayName { get; set; } = string.Empty;
        
        public bool IsActive { get; set; } = true;
        
        public bool IsCustom { get; set; } = false;

        public bool IsStrict { get; set; } = false;
        
        public int? UserId { get; set; }
    }
}
