using System;

namespace QicBoqMapper
{
    public class BoqRecord
    {
        public string Category { get; set; } = string.Empty;
        public string FamilyName { get; set; } = string.Empty;
        public string TypeName { get; set; } = string.Empty;
        public string ElementId { get; set; } = string.Empty;
        
        public string PackageNo { get; set; } = string.Empty;
        public string BillNo { get; set; } = string.Empty;
        public string SystemCode { get; set; } = string.Empty;
        public string PageNo { get; set; } = string.Empty;
        public string ItemNo { get; set; } = string.Empty;

        // Optional parameters
        public string DistrictCode { get; set; } = string.Empty;
        public string AssetGroup { get; set; } = string.Empty;
        public string AssetType { get; set; } = string.Empty;
        public string LocationCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string AbsL1 { get; set; } = string.Empty;
        public string AbsL2 { get; set; } = string.Empty;
        public string AbsL3 { get; set; } = string.Empty;
    }
}
