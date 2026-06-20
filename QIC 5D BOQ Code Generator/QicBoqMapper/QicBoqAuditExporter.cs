using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace QicBoqMapper
{
    public static class QicBoqAuditExporter
    {
        public static void ExportToExcel(string filePath, List<AuditRecord> records)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Audit Report");
                
                string[] headers = {
                    "Element Id", "Unique ID", "Category", "Family Name", "Type Name", 
                    "Package No", "Bill No", "System Code", "Page No", "Item No", 
                    "Generated BOQ Code", "Status", "Remarks"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[1, i + 1].Value = headers[i];
                    ws.Cells[1, i + 1].Style.Font.Bold = true;
                }

                int row = 2;
                foreach (var rec in records)
                {
                    ws.Cells[row, 1].Value = rec.ElementId;
                    ws.Cells[row, 2].Value = rec.UniqueId;
                    ws.Cells[row, 3].Value = rec.Category;
                    ws.Cells[row, 4].Value = rec.FamilyName;
                    ws.Cells[row, 5].Value = rec.TypeName;
                    ws.Cells[row, 6].Value = rec.PackageNo;
                    ws.Cells[row, 7].Value = rec.BillNo;
                    ws.Cells[row, 8].Value = rec.SystemCode;
                    ws.Cells[row, 9].Value = rec.PageNo;
                    ws.Cells[row, 10].Value = rec.ItemNo;
                    ws.Cells[row, 11].Value = rec.GeneratedBoqCode;
                    ws.Cells[row, 12].Value = rec.Status;
                    ws.Cells[row, 13].Value = rec.Remarks;
                    row++;
                }

                for (int i = 1; i <= headers.Length; i++)
                {
                    ws.Column(i).Width = 18;
                }
                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToCsv(string filePath, List<AuditRecord> records)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Element Id,Unique ID,Category,Family Name,Type Name,Package No,Bill No,System Code,Page No,Item No,Generated BOQ Code,Status,Remarks");

            foreach (var rec in records)
            {
                sb.AppendLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\"",
                    EscapeCsv(rec.ElementId),
                    EscapeCsv(rec.UniqueId),
                    EscapeCsv(rec.Category),
                    EscapeCsv(rec.FamilyName),
                    EscapeCsv(rec.TypeName),
                    EscapeCsv(rec.PackageNo),
                    EscapeCsv(rec.BillNo),
                    EscapeCsv(rec.SystemCode),
                    EscapeCsv(rec.PageNo),
                    EscapeCsv(rec.ItemNo),
                    EscapeCsv(rec.GeneratedBoqCode),
                    EscapeCsv(rec.Status),
                    EscapeCsv(rec.Remarks)
                ));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static void ExportToJson(string filePath, List<AuditRecord> records)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < records.Count; i++)
            {
                var rec = records[i];
                sb.AppendLine("  {");
                sb.AppendLine(string.Format("    \"ElementId\": \"{0}\",", EscapeJson(rec.ElementId)));
                sb.AppendLine(string.Format("    \"UniqueId\": \"{0}\",", EscapeJson(rec.UniqueId)));
                sb.AppendLine(string.Format("    \"Category\": \"{0}\",", EscapeJson(rec.Category)));
                sb.AppendLine(string.Format("    \"FamilyName\": \"{0}\",", EscapeJson(rec.FamilyName)));
                sb.AppendLine(string.Format("    \"TypeName\": \"{0}\",", EscapeJson(rec.TypeName)));
                sb.AppendLine(string.Format("    \"PackageNo\": \"{0}\",", EscapeJson(rec.PackageNo)));
                sb.AppendLine(string.Format("    \"BillNo\": \"{0}\",", EscapeJson(rec.BillNo)));
                sb.AppendLine(string.Format("    \"SystemCode\": \"{0}\",", EscapeJson(rec.SystemCode)));
                sb.AppendLine(string.Format("    \"PageNo\": \"{0}\",", EscapeJson(rec.PageNo)));
                sb.AppendLine(string.Format("    \"ItemNo\": \"{0}\",", EscapeJson(rec.ItemNo)));
                sb.AppendLine(string.Format("    \"GeneratedBoqCode\": \"{0}\",", EscapeJson(rec.GeneratedBoqCode)));
                sb.AppendLine(string.Format("    \"Status\": \"{0}\",", EscapeJson(rec.Status)));
                sb.AppendLine(string.Format("    \"Remarks\": \"{0}\"", EscapeJson(rec.Remarks)));
                sb.Append("  }" + (i < records.Count - 1 ? "," : "") + Environment.NewLine);
            }
            sb.AppendLine("]");
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static string EscapeCsv(string val)
        {
            if (val == null) return "";
            return val.Replace("\"", "\"\"");
        }

        private static string EscapeJson(string val)
        {
            if (val == null) return "";
            return val.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }
    }
}
