using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QicBoqMapper
{
    public static class QicBoqManager
    {
        public static List<Element> GetElements(Document doc, UIDocument uiDoc, string selectionMode, List<string> selectedCategoryNames, string categoryFilterText = "", string familyFilterText = "", string typeFilterText = "")
        {
            var elements = new List<Element>();
            if (selectedCategoryNames == null || selectedCategoryNames.Count == 0)
                return elements;

            var categoryFilters = new List<ElementFilter>();
            foreach (var name in selectedCategoryNames)
            {
                try
                {
                    var builtInCat = CategoryMappingService.GetBuiltInCategory(name);
                    categoryFilters.Add(new ElementCategoryFilter(builtInCat));
                }
                catch { }
            }

            if (categoryFilters.Count == 0)
                return elements;

            var categoryFilter = new LogicalOrFilter(categoryFilters);
            var collected = new List<Element>();

            if (selectionMode == "Current Selection" && uiDoc != null)
            {
                var selectionIds = uiDoc.Selection.GetElementIds();
                if (selectionIds.Count > 0)
                {
                    var collector = new FilteredElementCollector(doc, selectionIds);
                    collected.AddRange(collector.WherePasses(categoryFilter).WhereElementIsNotElementType().ToElements());
                }
            }
            else if (selectionMode == "Active View" && doc.ActiveView != null)
            {
                var collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                collected.AddRange(collector.WherePasses(categoryFilter).WhereElementIsNotElementType().ToElements());
            }
            else // Entire Model
            {
                var collector = new FilteredElementCollector(doc);
                collected.AddRange(collector.WherePasses(categoryFilter).WhereElementIsNotElementType().ToElements());
            }

            // Filter elements by category, family, and type name (case-insensitive substring match)
            foreach (var elem in collected)
            {
                if (!string.IsNullOrWhiteSpace(categoryFilterText))
                {
                    string catName = elem.Category?.Name ?? "";
                    if (catName.IndexOf(categoryFilterText, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                }

                var typeId = elem.GetTypeId();
                ElementType? typeElem = typeId != ElementId.InvalidElementId ? doc.GetElement(typeId) as ElementType : null;
                string familyName = typeElem?.FamilyName ?? "";
                string typeName = typeElem?.Name ?? elem.Name;

                if (!string.IsNullOrWhiteSpace(familyFilterText))
                {
                    if (familyName.IndexOf(familyFilterText, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                }

                if (!string.IsNullOrWhiteSpace(typeFilterText))
                {
                    if (typeName.IndexOf(typeFilterText, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                }

                elements.Add(elem);
            }

            return elements;
        }

        public static void CreateSharedParameters(Document doc, List<string> categoryNames)
        {
            var app = doc.Application;
            var categories = new CategorySet();

            foreach (var name in categoryNames)
            {
                try
                {
                    var builtIn = CategoryMappingService.GetBuiltInCategory(name);
                    var category = doc.Settings.Categories.get_Item(builtIn);
                    if (category != null && category.AllowsBoundParameters)
                    {
                        categories.Insert(category);
                    }
                }
                catch { }
            }

            if (categories.IsEmpty) return;

            // Define all required and optional parameters
            string[] paramNames = {
                "PACKAGE_NO", "BILL_NO", "SYSTEM_CODE", "PAGE_NO", "ITEM_NO", "QIC_5D_BOQ CODE",
                "DISTRICT_CODE", "ASSET_GROUP", "ASSET_TYPE", "LOCATION_CODE", "DESCRIPTION",
                "ABS_L1", "ABS_L2", "ABS_L3"
            };

            string originalFile = app.SharedParametersFilename;
            string tempFile = Path.Combine(Path.GetTempPath(), "QicTempSharedParameters.txt");

            try
            {
                using (StreamWriter sw = File.CreateText(tempFile))
                {
                    sw.WriteLine("# QIC 5D BOQ Shared Parameters");
                    sw.WriteLine("*META\tVERSION\tMINVERSION");
                    sw.WriteLine("META\t2\t1");
                    sw.WriteLine("*GROUP\tID\tNAME");
                    sw.WriteLine("GROUP\t1\tQIC Parameters");
                    sw.WriteLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE");
                    
                    // Generate GUIDs systematically
                    for (int i = 0; i < paramNames.Length; i++)
                    {
                        string guidVal = string.Format("d7a5e0f1-2244-4e99-bf12-7d5a6e0f{0:D4}", i + 1);
                        sw.WriteLine($"PARAM\t{guidVal}\t{paramNames[i]}\tTEXT\t\t1\t1\t\t1");
                    }
                }

                app.SharedParametersFilename = tempFile;
                var defFile = app.OpenSharedParameterFile();
                if (defFile == null) return;

                var group = defFile.Groups.get_Item("QIC Parameters");
                if (group == null) return;

                using (var t = new Transaction(doc, "Create QIC Shared Parameters"))
                {
                    t.Start();
                    var bindingMap = doc.ParameterBindings;

                    foreach (var pName in paramNames)
                    {
                        var def = group.Definitions.get_Item(pName);
                        if (def == null) continue;

                        bool alreadyBound = false;
                        var iterator = bindingMap.ForwardIterator();
                        while (iterator.MoveNext())
                        {
                            if (iterator.Key.Name == pName)
                            {
                                alreadyBound = true;
                                break;
                            }
                        }

                        if (!alreadyBound)
                        {
                            var binding = app.Create.NewInstanceBinding(categories);
                            BuiltInParameterGroup paramGroup = BuiltInParameterGroup.PG_DATA;
                            if (pName == "QIC_5D_BOQ CODE")
                            {
                                paramGroup = BuiltInParameterGroup.PG_IDENTITY_DATA;
                            }
                            else if (pName == "PACKAGE_NO" || pName == "BILL_NO" || pName == "SYSTEM_CODE" || 
                                     pName == "PAGE_NO" || pName == "ITEM_NO")
                            {
                                paramGroup = BuiltInParameterGroup.PG_TEXT;
                            }
                            bindingMap.Insert(def, binding, paramGroup);
                        }
                    }
                    t.Commit();
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(originalFile))
                    app.SharedParametersFilename = originalFile;
                try { File.Delete(tempFile); } catch { }
            }
        }

        public static List<AuditRecord> PerformMapping(Document doc, List<Element> elements, List<BoqRecord> excelRecords, string separator, string matchingMethod, bool caseInsensitive, out int matchedCount, out int unmatchedCount)
        {
            var auditRecords = new List<AuditRecord>();
            matchedCount = 0;
            unmatchedCount = 0;

            var comparer = caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

            // Element ID lookup (where ElementId column value is not empty)
            var idDict = excelRecords
                .Where(r => !string.IsNullOrEmpty(r.ElementId))
                .GroupBy(r => r.ElementId, comparer)
                .ToDictionary(g => g.Key, g => g.First(), comparer);

            // Category + Family + Type lookup
            var typeDict = excelRecords
                .Where(r => !string.IsNullOrEmpty(r.TypeName))
                .GroupBy(r => $"{r.Category}|{r.FamilyName}|{r.TypeName}", comparer)
                .ToDictionary(g => g.Key, g => g.First(), comparer);

            foreach (var element in elements)
            {
                var audit = new AuditRecord
                {
                    ElementId = element.Id.ToString(),
                    UniqueId = element.UniqueId,
                    Category = element.Category?.Name ?? "Unknown",
                };

                // Get type name and family
                var typeId = element.GetTypeId();
                ElementType typeElem = null;
                if (typeId != ElementId.InvalidElementId)
                {
                    typeElem = doc.GetElement(typeId) as ElementType;
                }

                string familyName = typeElem?.FamilyName ?? string.Empty;
                string typeName = typeElem?.Name ?? element.Name;
                audit.FamilyName = familyName;
                audit.TypeName = typeName;

                // Level
                var levelId = element.LevelId;
                var level = levelId != ElementId.InvalidElementId ? doc.GetElement(levelId) as Level : null;
                audit.Level = level?.Name ?? "";

                // Mark
                audit.Mark = element.LookupParameter("Mark")?.AsString() ?? "";

                // Workset
                try
                {
                    var ws = doc.GetWorksetTable()?.GetWorkset(element.WorksetId);
                    audit.Workset = ws?.Name ?? "";
                }
                catch { }

                // Matching
                BoqRecord matchedBoq = null;
                bool isMatched = false;

                // Match by Element ID first
                string elemIdStr = element.Id.ToString();
                if (idDict.TryGetValue(elemIdStr, out matchedBoq))
                {
                    isMatched = true;
                }
                // Fall back to Category + Family + Type Name
                else
                {
                    string key = $"{audit.Category}|{familyName}|{typeName}";
                    if (typeDict.TryGetValue(key, out matchedBoq))
                    {
                        isMatched = true;
                    }
                }

                if (isMatched && matchedBoq != null)
                {
                    audit.PackageNo = matchedBoq.PackageNo;
                    audit.BillNo = matchedBoq.BillNo;
                    audit.SystemCode = matchedBoq.SystemCode;
                    audit.PageNo = matchedBoq.PageNo;
                    audit.ItemNo = matchedBoq.ItemNo;

                    // Validate BOQ values
                    var missingFields = new List<string>();
                    if (string.IsNullOrEmpty(matchedBoq.PackageNo)) missingFields.Add("PACKAGE NO");
                    if (string.IsNullOrEmpty(matchedBoq.BillNo)) missingFields.Add("BILL NO");
                    if (string.IsNullOrEmpty(matchedBoq.SystemCode)) missingFields.Add("SYSTEM CODE");
                    if (string.IsNullOrEmpty(matchedBoq.PageNo)) missingFields.Add("PAGE NO");
                    if (string.IsNullOrEmpty(matchedBoq.ItemNo)) missingFields.Add("ITEM NO");

                    // Join fields using separator
                    string sepChar = "-";
                    if (separator == "Dot" || separator == "." || separator.Contains(".")) sepChar = ".";
                    else if (separator == "Underscore" || separator == "_" || separator.Contains("_")) sepChar = "_";
                    else if (separator == "Comma" || separator == "," || separator.Contains(",")) sepChar = ",";
                    else if (separator == "Dash" || separator == "-" || separator.Contains("-")) sepChar = "-";
                    else if (!string.IsNullOrEmpty(separator)) sepChar = separator;

                    audit.GeneratedBoqCode = string.Join(sepChar, new[] {
                        matchedBoq.PackageNo, matchedBoq.BillNo, matchedBoq.SystemCode, matchedBoq.PageNo, matchedBoq.ItemNo
                    }.Where(s => !string.IsNullOrEmpty(s)));

                    if (missingFields.Count > 0)
                    {
                        audit.Status = "Warning";
                        audit.Remarks = $"Matched, but missing values: {string.Join(", ", missingFields)}.";
                    }
                    else
                    {
                        audit.Status = "Success";
                        audit.Remarks = "Successfully matched.";
                    }
                    matchedCount++;
                }
                else
                {
                    audit.Status = "Error";
                    audit.Remarks = "No mapping found for this element type.";
                    unmatchedCount++;
                }

                auditRecords.Add(audit);
            }

            return auditRecords;
        }

        public static void UpdateParameters(Document doc, List<Element> elements, List<AuditRecord> auditRecords, List<BoqRecord> excelRecords, string matchingMethod, bool caseInsensitive)
        {
            var auditMap = auditRecords.ToDictionary(a => a.ElementId);
            var comparer = caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

            // Element ID lookup (where ElementId column value is not empty)
            var idDict = excelRecords
                .Where(r => !string.IsNullOrEmpty(r.ElementId))
                .GroupBy(r => r.ElementId, comparer)
                .ToDictionary(g => g.Key, g => g.First(), comparer);

            // Category + Family + Type lookup
            var typeDict = excelRecords
                .Where(r => !string.IsNullOrEmpty(r.TypeName))
                .GroupBy(r => $"{r.Category}|{r.FamilyName}|{r.TypeName}", comparer)
                .ToDictionary(g => g.Key, g => g.First(), comparer);

            using (var t = new Transaction(doc, "QIC 5D BOQ Update"))
            {
                t.Start();

                foreach (var element in elements)
                {
                    string elemIdStr = element.Id.ToString();
                    if (!auditMap.TryGetValue(elemIdStr, out var audit))
                        continue;

                    if (audit.Status == "Error")
                        continue;

                    // Set required instance parameters
                    SetParameterValue(element, "PACKAGE_NO", audit.PackageNo);
                    SetParameterValue(element, "BILL_NO", audit.BillNo);
                    SetParameterValue(element, "SYSTEM_CODE", audit.SystemCode);
                    SetParameterValue(element, "PAGE_NO", audit.PageNo);
                    SetParameterValue(element, "ITEM_NO", audit.ItemNo);
                    SetParameterValue(element, "QIC_5D_BOQ CODE", audit.GeneratedBoqCode);

                    // Lookup matching BoqRecord for optional parameters
                    BoqRecord matchedBoq = null;
                    bool found = false;

                    // Match by Element ID first
                    if (idDict.TryGetValue(elemIdStr, out matchedBoq))
                    {
                        found = true;
                    }
                    // Fall back to Category + Family + Type Name
                    else
                    {
                        var typeId = element.GetTypeId();
                        var typeElem = typeId != ElementId.InvalidElementId ? doc.GetElement(typeId) as ElementType : null;
                        string key = $"{audit.Category}|{typeElem?.FamilyName ?? ""}|{typeElem?.Name ?? element.Name}";
                        found = typeDict.TryGetValue(key, out matchedBoq);
                    }

                    if (found && matchedBoq != null)
                    {
                        SetParameterValue(element, "DISTRICT_CODE", matchedBoq.DistrictCode);
                        SetParameterValue(element, "ASSET_GROUP", matchedBoq.AssetGroup);
                        SetParameterValue(element, "ASSET_TYPE", matchedBoq.AssetType);
                        SetParameterValue(element, "LOCATION_CODE", matchedBoq.LocationCode);
                        SetParameterValue(element, "DESCRIPTION", matchedBoq.Description);
                        SetParameterValue(element, "ABS_L1", matchedBoq.AbsL1);
                        SetParameterValue(element, "ABS_L2", matchedBoq.AbsL2);
                        SetParameterValue(element, "ABS_L3", matchedBoq.AbsL3);
                    }
                }

                t.Commit();
            }
        }

        private static void SetParameterValue(Element element, string paramName, string value)
        {
            string[] names = GetParameterNameVariants(paramName);
            Parameter param = null;

            // Try lookup on Instance first
            foreach (var name in names)
            {
                param = element.LookupParameter(name);
                if (param != null) break;
            }

            // If not found on Instance, try Type
            if (param == null)
            {
                var typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    var typeElem = element.Document.GetElement(typeId);
                    if (typeElem != null)
                    {
                        foreach (var name in names)
                        {
                            param = typeElem.LookupParameter(name);
                            if (param != null) break;
                        }
                    }
                }
            }

            if (param != null && !param.IsReadOnly)
            {
                param.Set(value ?? string.Empty);
            }
        }

        private static string[] GetParameterNameVariants(string baseName)
        {
            var list = new List<string> { baseName };

            string spaceName = baseName.Replace("_", " ");
            if (!list.Contains(spaceName)) list.Add(spaceName);

            string titleSpace = System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(spaceName.ToLower());
            if (!list.Contains(titleSpace)) list.Add(titleSpace);

            string titleUnderscore = System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(baseName.ToLower());
            if (!list.Contains(titleUnderscore)) list.Add(titleUnderscore);

            string camelCase = titleSpace.Replace(" ", "");
            if (!list.Contains(camelCase)) list.Add(camelCase);

            string lowerName = baseName.ToLower();
            if (!list.Contains(lowerName)) list.Add(lowerName);

            return list.ToArray();
        }
    }
}
