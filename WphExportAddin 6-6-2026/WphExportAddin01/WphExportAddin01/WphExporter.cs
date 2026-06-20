using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using Autodesk.Revit.DB;

namespace WphExportAddin
{
    /// <summary>
    /// Core export engine. One public method (Run) orchestrates the formats
    /// according to the run mode. Each format is wrapped independently so a
    /// single failure cannot abort the rest of the run.
    ///
    /// LOOP + DELAY
    /// ────────────
    /// Run() now builds an ordered list of export operations (one per view per
    /// format) and runs them one at a time. If ExportSettings.DelaySeconds &gt; 0
    /// it pauses BETWEEN operations (not after the last). This is the small
    /// delay that reduces instability on big models.
    ///
    /// NWC STRATEGY — XML settings file (Export_NWC_setting.xml)
    /// ──────────────────────────────────────────────────────────
    /// The NWC export is driven by NavisworksExportOptions, resolved via
    /// reflection (so the add-in still loads even if the exporter is absent).
    /// ALL properties are mapped from the project's Export_NWC_setting.xml:
    ///
    ///   nwexportrevit_element_params            → Parameters         = All
    ///   nwexportrevit_element_ids               → ExportElementIds   = true
    ///   nwexportrevit_element_find_missing_materials → FindMissingMaterials = true
    ///   nwexportrevit_section_extract           → ExportScope (CurrentView)
    ///   nwexportrevit_coordinates               → Coordinates        = Shared
    ///   nwexportrevit_element_generic_properties→ ConvertElementProperties = true
    ///   nwexportrevit_urls                      → ExportUrls         = true
    ///   nwexportrevit_room                      → ExportRoomAsAttribute = true
    ///   nwexportrevit_room_geometry             → ExportRoomGeometry = false
    ///   nwexportrevit_linked_files              → ExportLinks        = false
    ///   nwexportrevit_construction_parts        → ExportParts        = false
    ///   nwexportrevit_divide_file_into_levels   → DivideFileIntoLevels = true
    ///   nwexportrevit_param_faceting_factor     → FacetingFactor     = 1.0
    ///   nwexportrevit_linked_CAD_formats        → ConvertLinkedCADFormats = false
    ///   nwexportrevit_lights                    → ConvertLights      = false
    ///   nwexportrevit_with_type_props           → (covered by ConvertElementProperties + Parameters=All)
    ///   nwexportrevit_separate_custom_props     → (no API member; Navisworks reader-side only)
    ///
    /// Property names (PascalCase) are the NavisworksExportOptions members in
    /// the Revit 2023+ API. Each assignment goes through TrySet* helpers so a
    /// missing/renamed member never crashes the export.
    /// </summary>
    public class WphExporter
    {
        private readonly Document _doc;
        private readonly ExportSettings _settings;
        private readonly StringBuilder _log = new StringBuilder();

        public WphExporter(Document doc, ExportSettings settings)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        private void Log(string line) => _log.AppendLine(line);

        // ─────────────────────────────────────────────────────────────────────
        // PUBLIC ENTRY POINT
        // ─────────────────────────────────────────────────────────────────────
        public string Run()
        {
            Log(string.Format("=== WPH EXPORT RUN @ {0} ===",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            Log(string.Format("mode={0}  modelName={1}  folder={2}  subfolders={3}  dateStamp={4}  delay={5}s  links={6}",
                _settings.Mode, _settings.ModelName, _settings.OutputFolder,
                _settings.UseSubfolders, _settings.DateStamp, _settings.DelaySeconds, _settings.Links));

            // Build (scope, viewName) passes – one per selected 3D view, or one
            // whole-model pass when no view is selected.
            var passes = new List<(ExportScope scope, string viewName)>();
            if (_settings.HasViewSelection)
            {
                foreach (var v in _settings.SelectedViews)
                    passes.Add((ExportScope.ForView(v), v.Name));
                Log(string.Format("scope=3D views ({0}): {1}",
                    _settings.SelectedViews.Count,
                    string.Join(", ", _settings.SelectedViews.ConvertAll(v => v.Name))));
            }
            else
            {
                passes.Add((ExportScope.BuildDefault(), null));
                Log("scope=(Whole model)");
            }

            // Flatten everything into one ordered list of export operations so
            // the delay can sit cleanly BETWEEN each item, whatever the mode.
            var ops = new List<(string label, Action run)>();
            switch (_settings.Mode)
            {
                case RunMode.Weekly:
                    foreach (var p in passes)
                    {
                        var pass = p;   // capture per iteration
                        ops.Add((Tag("NWC", pass.viewName),
                            () => ExportNwc(_doc, pass.scope, pass.viewName)));
                    }
                    break;

                case RunMode.Monthly:
                    foreach (var p in passes)
                    {
                        var pass = p;
                        ops.Add((Tag("NWC", pass.viewName),
                            () => ExportNwc(_doc, pass.scope, pass.viewName)));
                    }
                    foreach (var p in passes)
                    {
                        var pass = p;
                        ops.Add((Tag("IFC", pass.viewName),
                            () => ExportIfc(_doc, pass.scope, pass.viewName)));
                    }
                    break;
            }

            // SEPARATE-FILE LINK EXPORT: host already queued above WITHOUT links
            // (ExportLinks resolves to false for this mode); now add one whole-
            // model export per loaded Revit link, each as its own file.
            if (_settings.Links == LinkMode.SeparateFiles)
                AppendSeparateLinkOps(ops);

            // Run them one at a time, pausing between (never after the last).
            for (int i = 0; i < ops.Count; i++)
            {
                ops[i].run();

                if (i < ops.Count - 1 && _settings.DelaySeconds > 0)
                {
                    Log(string.Format("... waiting {0}s before next export ({1}) ...",
                        _settings.DelaySeconds, ops[i + 1].label));
                    Thread.Sleep(_settings.DelaySeconds * 1000);
                }
            }

            Log("=== RUN COMPLETE ===");
            Log("");
            Log("-----------------------------------------");
            Log(AddinInfo.CreditBlock());
            return _log.ToString();
        }

        private static string Tag(string fmt, string viewName)
            => fmt + (viewName == null ? " (whole model)" : " view: " + viewName);

        /// <summary>
        /// Queue one whole-model export per LOADED Revit link (deduplicated by
        /// the link's file path). Each link becomes its own file tagged
        /// "Link-&lt;name&gt;"; failures are isolated inside ExportNwc/ExportIfc.
        /// </summary>
        private void AppendSeparateLinkOps(List<(string label, Action run)> ops)
        {
            var linkDocs = GetUniqueLoadedLinkDocs();
            if (linkDocs.Count == 0)
            {
                Log("links=SeparateFiles -> no LOADED Revit links found in host (unloaded links are skipped).");
                return;
            }
            Log(string.Format("links=SeparateFiles -> {0} link(s) will be exported as their own files.",
                linkDocs.Count));

            foreach (var ld in linkDocs)
            {
                var linkDoc = ld;  // capture
                string name;
                try { name = Path.GetFileNameWithoutExtension(linkDoc.Title); }
                catch { name = "link"; }
                if (string.IsNullOrWhiteSpace(name)) name = "link";
                string tag = "Link-" + name;

                switch (_settings.Mode)
                {
                    case RunMode.Weekly:
                        ops.Add(("NWC " + tag,
                            () => ExportNwc(linkDoc, ExportScope.BuildDefault(), tag)));
                        break;
                    case RunMode.Monthly:
                        ops.Add(("NWC " + tag,
                            () => ExportNwc(linkDoc, ExportScope.BuildDefault(), tag)));
                        ops.Add(("IFC " + tag,
                            () => ExportIfc(linkDoc, ExportScope.BuildDefault(), tag)));
                        break;
                }
            }
        }

        /// <summary>
        /// Every loaded Revit link document in the host, deduplicated by file
        /// path (many RevitLinkInstances can point at the same document).
        /// Unloaded links return a null document and are skipped.
        /// </summary>
        private List<Document> GetUniqueLoadedLinkDocs()
        {
            var result = new List<Document>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var instances = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            foreach (var inst in instances)
            {
                Document ld = null;
                try { ld = inst.GetLinkDocument(); } catch { /* not loaded */ }
                if (ld == null) continue;

                string key = ld.PathName;
                if (string.IsNullOrEmpty(key)) key = ld.Title ?? Guid.NewGuid().ToString();
                if (seen.Add(key)) result.Add(ld);
            }
            return result;
        }

        // ═════════════════════════════════════════════════════════════════════
        // NWC — NavisworksExportOptions mapped from Export_NWC_setting.xml
        // ═════════════════════════════════════════════════════════════════════
        private void ExportNwc(Document doc, ExportScope scope, string viewName)
        {
            try
            {
                // ── 1. Resolve the NavisworksExportOptions type via reflection ──────
                Type navisType = ResolveNavisOptionsType();
                if (navisType == null)
                {
                    Log("NWC : SKIPPED  -> NavisworksExportOptions type not found in loaded assemblies. " +
                        "Install 'Navisworks Exporters' for your Revit version and restart Revit.");
                    return;
                }

                object opts = Activator.CreateInstance(navisType);

                // ── 2. SCOPE ──────────────────────────────────────────────────────
                // XML: nwexportrevit_section_extract = "Current view" (internal :1)
                // When the user has selected specific 3-D views, we honour them;
                // otherwise we export the whole model.
                if (scope.WholeModel)
                {
                    // ExportScope.Model  (enum value "Model")
                    TrySetEnum(navisType, opts, "ExportScope", "Model");
                }
                else
                {
                    // ExportScope.View  — export only the chosen 3D view.
                    TrySetEnum(navisType, opts, "ExportScope", "View");
                    TrySetProp(navisType, opts, "ViewId", scope.ViewId);
                }

                // ── 3. COORDINATES ───────────────────────────────────────────────
                // XML: nwexportrevit_coordinates = "Shared"  (internal :1)
                // NavisworksCoordinates enum: Shared
                TrySetEnum(navisType, opts, "Coordinates", "Shared");

                // ── 4. PARAMETERS ────────────────────────────────────────────────
                // XML: nwexportrevit_element_params = "All"  (internal :2)
                // NavisworksParameters enum: All
                TrySetEnum(navisType, opts, "Parameters", "All");

                // ── 5. ELEMENT IDs ───────────────────────────────────────────────
                // XML: nwexportrevit_element_ids = true
                TrySetBool(navisType, opts, "ExportElementIds", true);

                // ── 6. MISSING MATERIALS ─────────────────────────────────────────
                // XML: nwexportrevit_element_find_missing_materials = true
                TrySetBool(navisType, opts, "FindMissingMaterials", true);

                // ── 7. GENERIC PROPERTIES ────────────────────────────────────────
                // XML: nwexportrevit_element_generic_properties = true
                // Maps to ConvertElementProperties in NavisworksExportOptions.
                TrySetBool(navisType, opts, "ConvertElementProperties", true);

                // ── 8. TYPE PROPERTIES ───────────────────────────────────────────
                // XML: nwexportrevit_with_type_props = true
                // The NWC exporter has NO separate type-property toggle in the
                // API. Element + type parameters are both governed by
                // ConvertElementProperties (set above) together with
                // Parameters = All (set in step 4). Nothing further to set here.

                // ── 9. URLS / HYPERLINKS ─────────────────────────────────────────
                // XML: nwexportrevit_urls = true
                TrySetBool(navisType, opts, "ExportUrls", true);

                // ── 10. ROOMS ────────────────────────────────────────────────────
                // XML: nwexportrevit_room = true
                TrySetBool(navisType, opts, "ExportRoomAsAttribute", true);

                // XML: nwexportrevit_room_geometry = false
                TrySetBool(navisType, opts, "ExportRoomGeometry", false);

                // ── 11. LINKED FILES ─────────────────────────────────────────────
                // XML baseline: nwexportrevit_linked_files = false.
                // Now driven by the Links setting:
                //   IncludeInHost -> true  (links bundled into this NWC)
                //   None / SeparateFiles -> false (links handled elsewhere / not at all)
                bool includeLinks = _settings.Links == LinkMode.IncludeInHost;
                TrySetBool(navisType, opts, "ExportLinks", includeLinks);

                // XML: nwexportrevit_linked_CAD_formats = false
                // CORRECT API member is ConvertLinkedCADFormats. Its API default
                // is TRUE, so this assignment is required to honour the XML
                // (without it, linked CAD would be exported).
                TrySetBool(navisType, opts, "ConvertLinkedCADFormats", false);

                // ── 12. DIVIDE BY LEVELS ─────────────────────────────────────────
                // XML: nwexportrevit_divide_file_into_levels = true
                TrySetBool(navisType, opts, "DivideFileIntoLevels", true);

                // ── 13. CONSTRUCTION PARTS ───────────────────────────────────────
                // XML: nwexportrevit_construction_parts = false
                TrySetBool(navisType, opts, "ExportParts", false);

                // ── 14. FACETING FACTOR ──────────────────────────────────────────
                // XML: nwexportrevit_param_faceting_factor = 1.0  (full fidelity)
                TrySetDouble(navisType, opts, "FacetingFactor", 1.0);

                // ── 15. LIGHTS ───────────────────────────────────────────────────
                // XML: nwexportrevit_lights = false
                // CORRECT API member is ConvertLights (API default is false).
                TrySetBool(navisType, opts, "ConvertLights", false);

                // ── 16. SEPARATE CUSTOM PROPERTIES ──────────────────────────────
                // XML: nwexportrevit_separate_custom_props = true
                // NOTE: NavisworksExportOptions exposes NO member for this option
                // (it is a Navisworks-side reader preference, not part of the Revit
                // export API). There is nothing to set here; left as a no-op so the
                // mapping stays documented and complete.

                // ── 17. Resolve output paths ──────────────────────────────────────
                string dir = _settings.ResolveDir("nwc");
                string nameNoExt = Path.GetFileNameWithoutExtension(
                                       _settings.MakeFileName(doc.Title, "nwc", viewName));

                // ── 18. Find the correct Export overload via reflection ────────────
                // Document.Export(string folder, string name, NavisworksExportOptions)
                MethodInfo exportMethod = typeof(Document).GetMethod(
                    "Export",
                    new[] { typeof(string), typeof(string), navisType });

                if (exportMethod == null)
                {
                    Log("NWC : SKIPPED  -> Document.Export(string,string,NavisworksExportOptions) " +
                        "overload not found. The Navisworks Exporters add-in may be incomplete.");
                    return;
                }

                // ── 19. Log what we're about to do ────────────────────────────────
                string tag = viewName == null ? "(whole model)" : ("view: " + viewName);
                string linksNote = includeLinks ? "  links=ON" : "";
                Log(string.Format(
                    "NWC : STARTING -> dir={0}  file={1}.nwc  scope={2}  coords=Shared  params=All{3}",
                    dir, nameNoExt, tag, linksNote));

                // ── 20. Invoke the export ─────────────────────────────────────────
                // NavisworksExportOptions.Export returns void; success = no exception
                // + file exists on disk.
                exportMethod.Invoke(doc, new object[] { dir, nameNoExt, opts });

                string target = Path.Combine(dir, nameNoExt + ".nwc");
                if (File.Exists(target))
                {
                    var info = new FileInfo(target);
                    Log(string.Format("NWC : SUCCESS  -> {0}  ({1})  size={2:N0} KB",
                        target, tag, info.Length / 1024));
                }
                else
                {
                    Log("NWC : FAILED   -> Export call completed without error but output file " +
                        "was not found at: " + target +
                        "  (check Revit's own journal for the exporter's internal message).");
                }
            }
            catch (Exception ex)
            {
                Log("NWC : ERROR    -> " + Flatten(ex));
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        // IFC — IFCExportOptions + Document.Export
        // ═════════════════════════════════════════════════════════════════════
        private void ExportIfc(Document doc, ExportScope scope, string viewName)
        {
            try
            {
                IFCVersion ver = _settings.ResolveIfcVersion(out string verLabel);
                var opts = new IFCExportOptions { FileVersion = ver };

                if (scope.WholeModel || scope.ViewId == ElementId.InvalidElementId)
                {
                    opts.FilterViewId = ElementId.InvalidElementId;
                }
                else
                {
                    opts.FilterViewId = scope.ViewId;
                    try { opts.AddOption("ExportingView", "true"); } catch { }
                }

                opts.ExportBaseQuantities = true;
                opts.WallAndColumnSplitting = false;

                // LINKED FILES (best-effort). Whether the IFC exporter honours
                // this depends on the installed IFC exporter version, so we try
                // and log a note rather than promise federation. For guaranteed
                // link output use Links = SeparateFiles.
                if (_settings.Links == LinkMode.IncludeInHost && !doc.IsLinked)
                {
                    try { opts.AddOption("ExportLinkedFiles", "true"); } catch { }
                    Log("IFC : note     -> linked-files requested (IncludeInHost). Support is " +
                        "IFC-exporter-version dependent; use 'Separate files' for guaranteed link output.");
                }

                string dir = _settings.ResolveDir("ifc");
                string nameNoExt = Path.GetFileNameWithoutExtension(
                                       _settings.MakeFileName(doc.Title, "ifc", viewName));

                // IFC export of the HOST document is wrapped in a Transaction
                // (unchanged from the working single-model behaviour). A LINKED
                // document is read-only, so we must NOT open a transaction on it;
                // we export it directly and isolate any error.
                bool ok;
                if (doc.IsLinked)
                {
                    ok = doc.Export(dir, nameNoExt, opts);
                }
                else
                {
                    using (var tx = new Transaction(doc, "WPH IFC Export"))
                    {
                        tx.Start();
                        ok = doc.Export(dir, nameNoExt, opts);
                        tx.Commit();
                    }
                }

                string target = Path.Combine(dir, nameNoExt + ".ifc");
                string tag = viewName == null ? "(whole model)" : ("view: " + viewName);
                Log(ok
                    ? string.Format("IFC : SUCCESS  -> {0}  ({1};  scheme: {2})", target, tag, verLabel)
                    : string.Format("IFC : FAILED   -> Document.Export returned false.  target={0}", target));
            }
            catch (Exception ex)
            {
                Log("IFC : ERROR    -> " + Flatten(ex));
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        // Reflection helpers
        // ═════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Find NavisworksExportOptions in the loaded assemblies.
        /// Lives in RevitAPI.dll once the Navisworks Exporters add-in is installed.
        /// </summary>
        private static Type ResolveNavisOptionsType()
        {
            const string fullName = "Autodesk.Revit.DB.NavisworksExportOptions";

            // Fast path — same assembly as Document.
            Type t = typeof(Document).Assembly.GetType(fullName, throwOnError: false);
            if (t != null) return t;

            // Scan all loaded assemblies (exporter may be in a separate DLL).
            foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    t = asm.GetType(fullName, throwOnError: false);
                    if (t != null) return t;
                }
                catch { /* skip unreadable assemblies */ }
            }
            return null;
        }

        /// <summary>Set an enum property by its string member name (e.g. "Shared", "All").</summary>
        private static void TrySetEnum(Type t, object obj, string propName, string memberName)
        {
            try
            {
                PropertyInfo p = t.GetProperty(propName);
                if (p == null || !p.CanWrite) return;

                Type enumType = p.PropertyType;
                if (!enumType.IsEnum) return;

                if (Enum.IsDefined(enumType, memberName))
                    p.SetValue(obj, Enum.Parse(enumType, memberName));
            }
            catch { /* property absent or enum member renamed in this build */ }
        }

        /// <summary>Set a bool property.</summary>
        private static void TrySetBool(Type t, object obj, string propName, bool value)
        {
            try
            {
                PropertyInfo p = t.GetProperty(propName);
                if (p != null && p.PropertyType == typeof(bool) && p.CanWrite)
                    p.SetValue(obj, value);
            }
            catch { }
        }

        /// <summary>Set a double property (e.g. FacetingFactor).</summary>
        private static void TrySetDouble(Type t, object obj, string propName, double value)
        {
            try
            {
                PropertyInfo p = t.GetProperty(propName);
                if (p == null || !p.CanWrite) return;

                if (p.PropertyType == typeof(double))
                    p.SetValue(obj, value);
                else if (p.PropertyType == typeof(float))
                    p.SetValue(obj, (float)value);
            }
            catch { }
        }

        /// <summary>Set any property by value (used for ElementId, etc.).</summary>
        private static void TrySetProp(Type t, object obj, string propName, object value)
        {
            try
            {
                PropertyInfo p = t.GetProperty(propName);
                if (p != null && p.CanWrite)
                    p.SetValue(obj, value);
            }
            catch { }
        }

        /// <summary>
        /// Collapse an exception chain to one readable line.
        /// TargetInvocationException from reflection wraps the real error.
        /// </summary>
        private static string Flatten(Exception ex)
        {
            if (ex is TargetInvocationException tie && tie.InnerException != null)
                ex = tie.InnerException;
            return string.Format("{0}: {1}", ex.GetType().Name, ex.Message);
        }
    }
}
