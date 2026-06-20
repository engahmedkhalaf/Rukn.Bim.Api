using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using OfficeOpenXml;

namespace QicBoqMapper
{
    public static class QicBoqExportService
    {
        public static void ExportElements(string filePath, Document doc, List<Element> elements)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Revit Elements");
                
                string[] headers = {
                    "Element ID", "Unique ID", "Category", "Family Name", "Type Name",
                    "Level", "Workset", "Mark", "Package No", "Bill No",
                    "System Code", "Page No", "Item No", "QIC_5D_BOQ CODE"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[1, i + 1].Value = headers[i];
                    ws.Cells[1, i + 1].Style.Font.Bold = true;
                }

                int row = 2;
                foreach (var elem in elements)
                {
                    ws.Cells[row, 1].Value = elem.Id.ToString();
                    ws.Cells[row, 2].Value = elem.UniqueId;
                    ws.Cells[row, 3].Value = elem.Category?.Name ?? "Unknown";
                    
                    var typeId = elem.GetTypeId();
                    var typeElem = typeId != ElementId.InvalidElementId ? doc.GetElement(typeId) as ElementType : null;
                    ws.Cells[row, 4].Value = typeElem?.FamilyName ?? "";
                    ws.Cells[row, 5].Value = typeElem?.Name ?? elem.Name;

                    var levelId = elem.LevelId;
                    var level = levelId != ElementId.InvalidElementId ? doc.GetElement(levelId) as Level : null;
                    ws.Cells[row, 6].Value = level?.Name ?? "";

                    try
                    {
                        var worksetId = elem.WorksetId;
                        var workset = doc.GetWorksetTable()?.GetWorkset(worksetId);
                        ws.Cells[row, 7].Value = workset?.Name ?? "";
                    }
                    catch { ws.Cells[row, 7].Value = ""; }

                    ws.Cells[row, 8].Value = elem.LookupParameter("Mark")?.AsString() ?? "";

                    ws.Cells[row, 9].Value = elem.LookupParameter("PACKAGE_NO")?.AsString() ?? "";
                    ws.Cells[row, 10].Value = elem.LookupParameter("BILL_NO")?.AsString() ?? "";
                    ws.Cells[row, 11].Value = elem.LookupParameter("SYSTEM_CODE")?.AsString() ?? "";
                    ws.Cells[row, 12].Value = elem.LookupParameter("PAGE_NO")?.AsString() ?? "";
                    ws.Cells[row, 13].Value = elem.LookupParameter("ITEM_NO")?.AsString() ?? "";
                    ws.Cells[row, 14].Value = elem.LookupParameter("QIC_5D_BOQ CODE")?.AsString() ?? "";

                    row++;
                }

                var wsMap = package.Workbook.Worksheets.Add("BOQ Mapping");
                string[] mapHeaders = {
                    "Category", "Family Name", "Type Name", "Package No", "Bill No", "System Code", "Page No", "Item No",
                    "District Code", "Asset Group", "Asset Type", "Location Code", "Description", "ABS L1", "ABS L2", "ABS L3"
                };
                for (int i = 0; i < mapHeaders.Length; i++)
                {
                    wsMap.Cells[1, i + 1].Value = mapHeaders[i];
                    wsMap.Cells[1, i + 1].Style.Font.Bold = true;
                }

                for (int i = 1; i <= headers.Length; i++)
                {
                    ws.Column(i).Width = 18;
                }
                for (int i = 1; i <= mapHeaders.Length; i++)
                {
                    wsMap.Column(i).Width = 18;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}
