using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// License / Login command. Allows users to view their current activation status
    /// or activate a new license key at any time from the Revit ribbon tab.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphLoginCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            using (var licDlg = new LicenseDialog())
            {
                licDlg.ShowDialog();
            }
            return Result.Succeeded;
        }
    }
}
