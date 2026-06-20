using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// "About" command — shows the tool name, version, and the prepared-by /
    /// contact details (all sourced from AddinInfo so there's one place to edit).
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphAboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var td = new TaskDialog("About — " + AddinInfo.ToolName)
            {
                MainInstruction = AddinInfo.ToolName + "  v" + AddinInfo.Version,
                MainContent = AddinInfo.CreditBlock(),
                CommonButtons = TaskDialogCommonButtons.Close
            };
            td.Show();
            return Result.Succeeded;
        }
    }
}
