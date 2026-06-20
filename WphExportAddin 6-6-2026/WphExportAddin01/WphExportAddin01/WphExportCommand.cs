using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// Revit external command entry point. This is the class Revit instantiates
    /// when the user clicks the "WPH Export" button (wired via the .addin file).
    ///
    /// Transaction mode is Manual: the IFC/NWC export calls operate on the
    /// document directly and manage their own transactions internally, so we must
    /// NOT open an outer transaction.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphExportCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
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

            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument?.Document;

            if (doc == null)
            {
                message = "No active document. Open the WPH model first.";
                return Result.Failed;
            }

            // A SaveAs on an unsaved (never-saved) document needs a path; that's
            // fine since we always SaveAs to a NEW deliverable path. But IFC/NWC
            // export of a brand-new unsaved model can misbehave, so warn gently.
            try
            {
                // 1) Gather settings from the user via the simple dialog.
                var settings = new ExportSettings();
                using (var dlg = new InputDialog(settings, doc))
                {
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return Result.Cancelled; // user closed/cancelled
                    // dlg has written user choices back into `settings`.
                }

                // 2) Run the export. Each format is internally error-isolated.
                var exporter = new WphExporter(doc, settings);
                string log = exporter.Run();

                // 3) Show the log to the user (and it's easy to also write it to
                //    a .txt next to the deliverables if you want — see note below).
                TaskDialog td = new TaskDialog("WPH Export — Result")
                {
                    MainInstruction = "Export run complete.",
                    MainContent = log,
                    CommonButtons = TaskDialogCommonButtons.Close
                };
                td.Show();

                // OPTIONAL: persist the log to disk alongside deliverables:
                // System.IO.File.WriteAllText(
                //     System.IO.Path.Combine(settings.OutputFolder,
                //         "WPH-ExportLog-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt"),
                //     log);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // This catches only setup/dialog errors; per-format failures are
                // already handled inside WphExporter and reported in the log.
                message = "WPH Export failed before running: " + ex.Message;
                return Result.Failed;
            }
        }
    }
}
