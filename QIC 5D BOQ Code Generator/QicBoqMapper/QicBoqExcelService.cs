using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace QicBoqMapper
{
    public static class QicBoqExcelService
    {
        public static List<BoqRecord> LoadBoqRecords(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var records = new List<BoqRecord>();

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["Revit Elements"]
                                ?? package.Workbook.Worksheets["BOQ Mapping"]
                                ?? (package.Workbook.Worksheets.Count > 0 ? package.Workbook.Worksheets[0] : null);

                if (worksheet == null)
                    throw new InvalidOperationException("Could not find a worksheet to read mapping data.");

                int categoryCol = -1;
                int familyNameCol = -1;
                int typeNameCol = -1;
                int elementIdCol = -1;
                int packageNoCol = -1;
                int billNoCol = -1;
                int systemCodeCol = -1;
                int pageNoCol = -1;
                int itemNoCol = -1;

                int districtCodeCol = -1;
                int assetGroupCol = -1;
                int assetTypeCol = -1;
                int locationCodeCol = -1;
                int descCol = -1;
                int absL1Col = -1;
                int absL2Col = -1;
                int absL3Col = -1;

                int maxCols = 200;
                int emptyHeaderCount = 0;
                for (int col = 1; col <= maxCols; col++)
                {
                    string header = worksheet.Cells[1, col].Text?.Trim() ?? "";
                    if (string.IsNullOrEmpty(header))
                    {
                        emptyHeaderCount++;
                        if (emptyHeaderCount >= 10) break;
                        continue;
                    }
                    emptyHeaderCount = 0;

                    string headerUpper = header.ToUpper();
                    if (headerUpper == "CATEGORY") categoryCol = col;
                    else if (headerUpper == "FAMILY NAME" || headerUpper == "FAMILY") familyNameCol = col;
                    else if (headerUpper == "TYPE NAME" || headerUpper == "TYPE") typeNameCol = col;
                    else if (headerUpper == "ELEMENT ID" || headerUpper == "ELEMENTID") elementIdCol = col;
                    else if (headerUpper == "PACKAGE NO" || headerUpper == "PACKAGE_NO") packageNoCol = col;
                    else if (headerUpper == "BILL NO" || headerUpper == "BILL_NO") billNoCol = col;
                    else if (headerUpper == "SYSTEM CODE" || headerUpper == "SYSTEM_CODE") systemCodeCol = col;
                    else if (headerUpper == "PAGE NO" || headerUpper == "PAGE_NO") pageNoCol = col;
                    else if (headerUpper == "ITEM NO" || headerUpper == "ITEM_NO") itemNoCol = col;
                    else if (headerUpper == "DISTRICT CODE" || headerUpper == "DISTRICT_CODE") districtCodeCol = col;
                    else if (headerUpper == "ASSET GROUP" || headerUpper == "ASSET_GROUP") assetGroupCol = col;
                    else if (headerUpper == "ASSET TYPE" || headerUpper == "ASSET_TYPE") assetTypeCol = col;
                    else if (headerUpper == "LOCATION CODE" || headerUpper == "LOCATION_CODE") locationCodeCol = col;
                    else if (headerUpper == "DESCRIPTION") descCol = col;
                    else if (headerUpper == "ABS L1" || headerUpper == "ABS_L1") absL1Col = col;
                    else if (headerUpper == "ABS L2" || headerUpper == "ABS_L2") absL2Col = col;
                    else if (headerUpper == "ABS L3" || headerUpper == "ABS_L3") absL3Col = col;
                }

                var missing = new List<string>();
                if (typeNameCol == -1 && elementIdCol == -1) missing.Add("TYPE NAME or ELEMENT ID");
                if (packageNoCol == -1) missing.Add("PACKAGE NO");
                if (billNoCol == -1) missing.Add("BILL NO");
                if (systemCodeCol == -1) missing.Add("SYSTEM CODE");
                if (pageNoCol == -1) missing.Add("PAGE NO");
                if (itemNoCol == -1) missing.Add("ITEM NO");

                if (missing.Count > 0)
                {
                    throw new InvalidOperationException($"Required columns are missing: {string.Join(", ", missing)}");
                }

                int emptyStreak = 0;
                int maxRows = 100000;
                for (int row = 2; row <= maxRows; row++)
                {
                    string typeName = typeNameCol != -1 ? worksheet.Cells[row, typeNameCol].Text?.Trim() ?? "" : "";
                    string elementId = elementIdCol != -1 ? worksheet.Cells[row, elementIdCol].Text?.Trim() ?? "" : "";

                    if (string.IsNullOrEmpty(typeName) && string.IsNullOrEmpty(elementId))
                    {
                        emptyStreak++;
                        if (emptyStreak >= 20) break;
                        continue;
                    }
                    emptyStreak = 0;

                    records.Add(new BoqRecord
                    {
                        Category = categoryCol != -1 ? worksheet.Cells[row, categoryCol].Text?.Trim() ?? "" : "",
                        FamilyName = familyNameCol != -1 ? worksheet.Cells[row, familyNameCol].Text?.Trim() ?? "" : "",
                        TypeName = typeName,
                        ElementId = elementId,
                        PackageNo = worksheet.Cells[row, packageNoCol].Text?.Trim() ?? "",
                        BillNo = worksheet.Cells[row, billNoCol].Text?.Trim() ?? "",
                        SystemCode = worksheet.Cells[row, systemCodeCol].Text?.Trim() ?? "",
                        PageNo = worksheet.Cells[row, pageNoCol].Text?.Trim() ?? "",
                        ItemNo = worksheet.Cells[row, itemNoCol].Text?.Trim() ?? "",
                        
                        DistrictCode = districtCodeCol != -1 ? worksheet.Cells[row, districtCodeCol].Text?.Trim() ?? "" : "",
                        AssetGroup = assetGroupCol != -1 ? worksheet.Cells[row, assetGroupCol].Text?.Trim() ?? "" : "",
                        AssetType = assetTypeCol != -1 ? worksheet.Cells[row, assetTypeCol].Text?.Trim() ?? "" : "",
                        LocationCode = locationCodeCol != -1 ? worksheet.Cells[row, locationCodeCol].Text?.Trim() ?? "" : "",
                        Description = descCol != -1 ? worksheet.Cells[row, descCol].Text?.Trim() ?? "" : "",
                        AbsL1 = absL1Col != -1 ? worksheet.Cells[row, absL1Col].Text?.Trim() ?? "" : "",
                        AbsL2 = absL2Col != -1 ? worksheet.Cells[row, absL2Col].Text?.Trim() ?? "" : "",
                        AbsL3 = absL3Col != -1 ? worksheet.Cells[row, absL3Col].Text?.Trim() ?? "" : ""
                    });
                }
            }

            return records;
        }
    }
}
