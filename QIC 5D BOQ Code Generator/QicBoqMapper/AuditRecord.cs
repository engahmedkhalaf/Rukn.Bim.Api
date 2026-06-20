using System;

namespace QicBoqMapper
{
    public class AuditRecord
    {
        public string ElementId { get; set; } = string.Empty;
        public string UniqueId { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string FamilyName { get; set; } = string.Empty;
        public string TypeName { get; set; } = string.Empty;
        public string Level { get; set; } = string.Empty;
        public string Workset { get; set; } = string.Empty;
        public string Mark { get; set; } = string.Empty;

        public string PackageNo { get; set; } = string.Empty;
        public string BillNo { get; set; } = string.Empty;
        public string SystemCode { get; set; } = string.Empty;
        public string PageNo { get; set; } = string.Empty;
        public string ItemNo { get; set; } = string.Empty;
        public string GeneratedBoqCode { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty; // Success, Warning, Error
        public string Remarks { get; set; } = string.Empty;
    }
}
