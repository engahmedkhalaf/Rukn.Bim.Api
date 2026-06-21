using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;

namespace WphExportAddin
{
    /// <summary>
    /// Single place for the add-in's authorship/contact details. Edit here if
    /// anything changes; the ribbon About button and the result-log footer both
    /// read from these constants.
    /// </summary>
    public static class AddinInfo
    {
        public const string ToolName  = "RUKN BIM API Export";
        public const string Version   = "2.0";
        public const string PreparedBy = "Ahmed Khalaf";
        public const string Title      = "BIM Manager";
        public const string Company    = "BUJV";
        public const string Email      = "engkhalaf7@gmail.com";
        public const string Phone      = "0542554127";

        /// <summary>Multi-line credit block used in dialogs and the result log.</summary>
        public static string CreditBlock()
        {
            return
                "Prepared by: " + PreparedBy + " - " + Title + ", " + Company + "\n" +
                "Email: " + Email + "\n" +
                "Contact: " + Phone;
        }
    }

    /// <summary>
    /// Run cadence. weekly = NWC only; monthly = NWC + IFC.
    /// </summary>
    public enum RunMode
    {
        Weekly,
        Monthly
    }

    /// <summary>
    /// How Revit links are handled during export.
    ///   None          : links are NOT exported (NWC ExportLinks=false). Default.
    ///   IncludeInHost : bundle links into the host file. NWC sets ExportLinks=true;
    ///                   IFC requests linked-file inclusion best-effort (support is
    ///                   IFC-exporter-version dependent — see the log note).
    ///   SeparateFiles : host is exported WITHOUT links, then every loaded Revit
    ///                   link is exported as its OWN file (whole-link, deduped,
    ///                   error-isolated). This is the reliable, append-in-Navisworks
    ///                   workflow and the most predictable for IFC too.
    /// </summary>
    public enum LinkMode
    {
        None,
        IncludeInHost,
        SeparateFiles
    }

    /// <summary>
    /// A single 3D view the user can choose to export. Pairs the view's display
    /// name (for the checkbox label and for building unique filenames) with its
    /// ElementId (for the actual export call).
    /// </summary>
    public class ViewChoice
    {
        public string Name { get; set; }
        public ElementId Id { get; set; }
        public override string ToString() => Name;
    }

    /// <summary>
    /// All user-configurable export parameters in one place. This is the C#
    /// equivalent of the Dynamo IN[] ports. The InputDialog populates an
    /// instance of this; WphExporter consumes it.
    ///
    /// Filename pattern (fixed by brief):  WPH-&lt;DISCIPLINE&gt;-Model[.-date].&lt;ext&gt;
    /// </summary>
    public class ExportSettings
    {
        // -------- user inputs --------
        public RunMode Mode { get; set; } = RunMode.Weekly;
        public string OutputFolder { get; set; } = @"C:\RuknBIM\Deliverables";
        public string ModelName { get; set; } = "";
        public string IfcVersionKey { get; set; } = "IFC2x3CV2";
        public bool DateStamp { get; set; } = false;        // easy on/off
        public bool UseSubfolders { get; set; } = false;    // single folder by default

        // -------- linked-model handling --------
        public LinkMode Links { get; set; } = LinkMode.None;

        // -------- pacing --------
        // Seconds to pause BETWEEN export operations. Applied in two places so a
        // single setting controls every level of the loop:
        //   - WphExporter : between each view/format export inside one model
        //   - BatchExporter : between each model file
        // 0 = no delay (default). Raising this (e.g. 3-10s) reduces instability
        // on large models by giving Revit time to release handles / settle.
        public int DelaySeconds { get; set; } = 0;

        // -------- 3D view selection (MULTI-SELECT) --------
        // The user ticks one or more 3D views in the dialog's checkbox list.
        // BOTH NWC and IFC use these views (one output file per view), so the
        // two deliverables stay on the same scope.
        //   - SelectedViews empty  => "(Whole model)": one whole-model export
        //   - SelectedViews has N  => N exports, one per ticked view
        // Both formats (NWC/IFC) follow these same selected views.
        public List<ViewChoice> SelectedViews { get; set; } = new List<ViewChoice>();

        /// <summary>True when the user ticked at least one real 3D view.</summary>
        public bool HasViewSelection => SelectedViews != null && SelectedViews.Count > 0;

        // -------- fixed project constant --------
        public const string Prefix = "RUKN";

        // -----------------------------------------------------------------
        // Per-format subfolder mapping. Only used when UseSubfolders == true.
        // TO CHANGE: edit the values here (e.g. "Navisworks" instead of "NWC").
        // -----------------------------------------------------------------
        private static readonly Dictionary<string, string> SubfolderMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "nwc", "NWC" },
                { "ifc", "IFC" },
            };

        /// <summary>
        /// Build the deliverable filename for a given extension, optionally
        /// tagged with a view name so multi-view exports don't overwrite.
        /// e.g. ("MyRevitModel", "ifc")                -> "MyRevitModel.ifc"
        ///      ("MyRevitModel", "ifc", "Coordination")-> "MyRevitModel-Coordination.ifc"
        ///      (+ DateStamp)                          -> "MyRevitModel-Coordination-2026-05-25.ifc"
        /// </summary>
        public string MakeFileName(string defaultDocTitle, string ext, string viewName = null)
        {
            ext = ext.TrimStart('.');

            // Use custom ModelName if provided; otherwise, fallback to the document title.
            string targetTitle = !string.IsNullOrWhiteSpace(ModelName) && ModelName != "<Use Model Name>"
                ? ModelName 
                : defaultDocTitle;

            // Clean/sanitize the Revit document title
            string cleanDocTitle = SanitizeForFile(targetTitle);
            
            string baseName = cleanDocTitle;

            // Append the view name (sanitised) when exporting a specific view.
            if (!string.IsNullOrWhiteSpace(viewName))
                baseName = string.Format("{0}-{1}", baseName, SanitizeForFile(viewName));

            if (DateStamp)
            {
                string stamp = DateTime.Now.ToString("yyyy-MM-dd");
                baseName = string.Format("{0}-{1}", baseName, stamp);
            }
            return string.Format("{0}.{1}", baseName, ext);
        }

        /// <summary>Strip characters Windows won't allow in a filename.</summary>
        private static string SanitizeForFile(string raw)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                raw = raw.Replace(c, '_');
            // Also collapse spaces to keep names tidy.
            return raw.Replace(' ', '_').Trim('_');
        }

        /// <summary>
        /// Resolve (and create if missing) the output directory for a format.
        /// Default: single OutputFolder. With UseSubfolders: OutputFolder\&lt;FMT&gt;.
        /// </summary>
        public string ResolveDir(string ext)
        {
            string target = OutputFolder;
            if (UseSubfolders)
            {
                string sub = SubfolderMap.ContainsKey(ext)
                    ? SubfolderMap[ext]
                    : ext.ToUpperInvariant();
                target = Path.Combine(OutputFolder, sub);
            }

            if (!Directory.Exists(target))
                Directory.CreateDirectory(target); // throws on failure -> caught per-format

            return target;
        }

        /// <summary>Full path = ResolveDir + MakeFileName (optional view tag).</summary>
        public string FullPath(string docTitle, string ext, string viewName = null)
        {
            return Path.Combine(ResolveDir(ext), MakeFileName(docTitle, ext, viewName));
        }

        // -----------------------------------------------------------------
        // IFC version map. Default IFC2x3 Coordination View 2.0.
        // Returns the enum plus a human label (notes fallback if key unknown).
        // -----------------------------------------------------------------
        private static readonly Dictionary<string, IFCVersion> IfcVersionMap =
            new Dictionary<string, IFCVersion>(StringComparer.OrdinalIgnoreCase)
            {
                { "IFC2x3CV2", IFCVersion.IFC2x3CV2 },  // DEFAULT
                { "IFC2x3",    IFCVersion.IFC2x3 },
                { "IFC2x2",    IFCVersion.IFC2x2 },
                { "IFC4",      IFCVersion.IFC4 },
                { "IFC4RV",    IFCVersion.IFC4RV },
                { "IFC4DTV",   IFCVersion.IFC4DTV },
                { "IFCBCA",    IFCVersion.IFCBCA },
                { "IFCCOBIE",  IFCVersion.IFCCOBIE }
            };

        /// <summary>Available IFC keys, for the dropdown in the dialog.</summary>
        public static IEnumerable<string> IfcVersionKeys => IfcVersionMap.Keys;

        public IFCVersion ResolveIfcVersion(out string label)
        {
            if (IfcVersionMap.TryGetValue(IfcVersionKey, out IFCVersion v))
            {
                label = IfcVersionKey;
                return v;
            }
            label = string.Format("IFC2x3CV2 (fallback; '{0}' unknown)", IfcVersionKey);
            return IFCVersion.IFC2x3CV2;
        }
    }

    /// <summary>
    /// Shared export scope - the single source of truth for what NWC and IFC
    /// cover. Both formats read from this so they stay consistent.
    ///
    /// UPDATED: the scope now comes from the user's chosen 3D view.
    ///   - WholeModel == true   => export entire project (view ignored)
    ///   - WholeModel == false  => export only ViewId (a 3D view)
    /// </summary>
    public class ExportScope
    {
        public bool WholeModel { get; set; } = true;
        public ElementId ViewId { get; set; } = ElementId.InvalidElementId;

        /// <summary>Whole-model scope (no view filter). Kept for fallback/default.</summary>
        public static ExportScope BuildDefault()
        {
            return new ExportScope
            {
                WholeModel = true,
                ViewId = ElementId.InvalidElementId
            };
        }

        /// <summary>
        /// Build a scope for one specific view (used when looping selected views).
        /// </summary>
        public static ExportScope ForView(ViewChoice v)
        {
            bool hasView = v != null && v.Id != null && v.Id != ElementId.InvalidElementId;
            return new ExportScope
            {
                WholeModel = !hasView,
                ViewId = hasView ? v.Id : ElementId.InvalidElementId
            };
        }
    }
}
