using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// Shared helpers for the one-click commands. Both Weekly and Monthly use the
    /// same flow: build sensible defaults, prompt ONLY for an output folder, then
    /// run the exporter and show the result log.
    /// </summary>
    internal static class OneClickHelper
    {
        public static Result Run(ExternalCommandData commandData, RunMode mode, ref string message)
        {
            // 1) Verify license status first
            if (!LicenseManager.IsLicensed())
            {
                using (var licDlg = new LicenseDialog())
                {
                    if (licDlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        message = "WPH Export requires a valid active license.";
                        return Result.Cancelled;
                    }
                }
            }

            Document doc = commandData.Application.ActiveUIDocument?.Document;
            if (doc == null)
            {
                message = "No active document. Open the WPH model first.";
                return Result.Failed;
            }

            // Sensible defaults for one-click runs. Edit here if you want a fixed
            // discipline / folder / IFC version baked in for a true zero-click run.
            var settings = new ExportSettings
            {
                Mode = mode,
                ModelName = "",
                IfcVersionKey = "IFC2x3CV2",
                DateStamp = false,
                UseSubfolders = false,
                SelectedViews = new System.Collections.Generic.List<ViewChoice>()  // empty = whole model
            };

            // Ask ONLY for the output folder so the run still goes somewhere sane.
            // To make it truly one-click (no prompts), comment this block out and
            // hardcode settings.OutputFolder above.
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Choose the output folder for the " + mode + " export";
                if (!string.IsNullOrWhiteSpace(settings.OutputFolder) &&
                    Directory.Exists(settings.OutputFolder))
                    fbd.SelectedPath = settings.OutputFolder;

                if (fbd.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;
                settings.OutputFolder = fbd.SelectedPath;
            }

            try
            {
                var exporter = new WphExporter(doc, settings);
                string log = exporter.Run();

                var td = new TaskDialog("WPH Export — " + mode + " result")
                {
                    MainInstruction = mode + " export complete.",
                    MainContent = log,
                    CommonButtons = TaskDialogCommonButtons.Close
                };
                td.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "WPH " + mode + " export failed: " + ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>One-click weekly export (NWC only).</summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphExportWeeklyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            => OneClickHelper.Run(commandData, RunMode.Weekly, ref message);
    }

    /// <summary>One-click monthly export (NWC + IFC).</summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphExportMonthlyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            => OneClickHelper.Run(commandData, RunMode.Monthly, ref message);
    }
}
