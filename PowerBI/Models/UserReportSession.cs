using System;
using System.ComponentModel.DataAnnotations;

namespace PowerBI.Models
{
    public class UserReportSession
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int UserId { get; set; }

        [Required]
        public int ReportId { get; set; }

        [Required]
        public string TempPowerBIReportId { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime ExpiresAt { get; set; } = DateTime.UtcNow.AddHours(1);

        public bool IsActive { get; set; } = true;
        
        public string? FilterHash { get; set; }
    }
}
