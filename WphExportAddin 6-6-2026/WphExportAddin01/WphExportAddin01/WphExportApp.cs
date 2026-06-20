using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// IExternalApplication = startup/shutdown hook for the add-in. Revit calls
    /// OnStartup() once when Revit launches; this is where we build our ribbon
    /// tab, panel, and the SplitButton (main button + dropdown items).
    ///
    /// Wired in WphExportAddin.addin via &lt;AddIn Type="Application"&gt;.
    /// </summary>
    public class WphExportApp : IExternalApplication
    {
        // The tab and panel names users see in Revit.
        private const string TabName = "WPH";
        private const string PanelName = "Export";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // 1) Create a dedicated ribbon TAB. Wrap in try/catch because
                //    Revit throws if the tab already exists (e.g. another add-in
                //    or a prior load created it).
                try { application.CreateRibbonTab(TabName); } catch { /* tab exists */ }

                // 2) Create or retrieve a PANEL inside that tab safely.
                RibbonPanel panel = null;
                try
                {
                    var existingPanels = application.GetRibbonPanels(TabName);
                    foreach (var p in existingPanels)
                    {
                        if (p.Name.Equals(PanelName, StringComparison.OrdinalIgnoreCase))
                        {
                            panel = p;
                            break;
                        }
                    }
                }
                catch { }

                if (panel == null)
                {
                    panel = application.CreateRibbonPanel(TabName, PanelName);
                }

                // 3) Resolve where THIS DLL lives, so we can reference its types
                //    and find the icon resources copied alongside it.
                string asmPath = Assembly.GetExecutingAssembly().Location;

                // 4) Build the three command "definitions" that the SplitButton
                //    needs. Each one targets a class implementing IExternalCommand.
                var mainBtnData = new PushButtonData(
                    "WphExportMain",                     // internal name (unique)
                    "WPH Export",                        // label under the icon
                    asmPath,
                    "WphExportAddin.WphExportCommand")   // full class name
                {
                    ToolTip = "Open the WPH export dialog (pick scope, format options, run).",
                    LongDescription = "Click the main button to open the WPH export dialog. " +
                                      "Click the arrow below for one-click Weekly or Monthly runs."
                };

                var weeklyBtnData = new PushButtonData(
                    "WphExportWeekly",
                    "Weekly (NWC)",
                    asmPath,
                    "WphExportAddin.WphExportWeeklyCommand")
                {
                    ToolTip = "One-click weekly export: NWC only, whole model, default settings.",
                };

                var monthlyBtnData = new PushButtonData(
                    "WphExportMonthly",
                    "Monthly (NWC+IFC)",
                    asmPath,
                    "WphExportAddin.WphExportMonthlyCommand")
                {
                    ToolTip = "One-click monthly export: NWC + IFC, default settings.",
                };

                // Batch: pick MANY .rvt files and export them one after the
                // other with a delay between each. Whole-model scope per file.
                var batchBtnData = new PushButtonData(
                    "WphExportBatch",
                    "Batch (multiple models)",
                    asmPath,
                    "WphExportAddin.WphBatchExportCommand")
                {
                    ToolTip = "Export several Revit models in a row, with a pause between each.",
                    LongDescription = "Pick multiple .rvt files. Each is opened, exported " +
                                      "(same NWC/IFC settings as the main command), then closed. " +
                                      "A configurable delay is inserted between models and every " +
                                      "model/format is error-isolated so one failure can't stop the run."
                };

                // 5) Attach icons to each definition.
                AttachIcons(mainBtnData, "WphExport");
                AttachIcons(weeklyBtnData, "WphExportWeekly");
                AttachIcons(monthlyBtnData, "WphExportMonthly");
                AttachIcons(batchBtnData, "WphExportBatch");

                // 6) Create the SplitButton. The "current" button (main face) is
                //    the dialog launcher; the dropdown adds the two one-click runs.
                var splitData = new SplitButtonData("WphExportSplit", "WPH Export");
                SplitButton split = panel.AddItem(splitData) as SplitButton;

                PushButton mainBtn = split.AddPushButton(mainBtnData);
                split.AddPushButton(weeklyBtnData);
                split.AddPushButton(monthlyBtnData);
                split.AddPushButton(batchBtnData);

                // Show the dialog launcher as the default face of the SplitButton.
                split.CurrentButton = mainBtn;

                // Separator + an "About" and "License" buttons.
                panel.AddSeparator();

                var loginData = new PushButtonData(
                    "WphLogin",
                    "License",
                    asmPath,
                    "WphExportAddin.WphLoginCommand")
                {
                    ToolTip = "View or activate your WPH Export license."
                };
                AttachIcons(loginData, "WphLicense");
                panel.AddItem(loginData);

                var aboutData = new PushButtonData(
                    "WphAbout",
                    "About",
                    asmPath,
                    "WphExportAddin.WphAboutCommand")
                {
                    ToolTip = "Tool info and contact details.",
                    LongDescription = AddinInfo.CreditBlock()
                };
                AttachIcons(aboutData, "WphAbout");
                panel.AddItem(aboutData);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Don't kill Revit's startup if ribbon creation fails; just log.
                TaskDialog.Show("WPH Export — startup error",
                    "Could not build the WPH ribbon:\n\n" + ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Load the WPH icons from a Resources folder placed next to the DLL.
        /// If a button-specific icon is not found, falls back to the default WphExport icon.
        /// </summary>
        private static void AttachIcons(PushButtonData btn, string baseName)
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            string icon32 = Path.Combine(dir, "Resources", baseName + "_32.png");
            string icon16 = Path.Combine(dir, "Resources", baseName + "_16.png");

            // Fallback to default icons if specific files don't exist
            if (!File.Exists(icon32))
                icon32 = Path.Combine(dir, "Resources", "WphExport_32.png");
            if (!File.Exists(icon16))
                icon16 = Path.Combine(dir, "Resources", "WphExport_16.png");

            if (File.Exists(icon32))
                btn.LargeImage = new BitmapImage(new Uri(icon32, UriKind.Absolute));
            if (File.Exists(icon16))
                btn.Image = new BitmapImage(new Uri(icon16, UriKind.Absolute));
        }
    }
}
