using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QicBoqMapper
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QicBoqManagerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, null);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExportElementsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "export");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ImportMappingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "import");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Generate5DCodesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "generate");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ValidateMappingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "validate");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AuditReportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "audit");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SettingsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!CommandHelper.EnsureActivated(commandData.Application))
                {
                    return Result.Cancelled;
                }

                QicBoqApp.ShowWindow(commandData.Application, "settings");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LicenseCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var win = new LicenseWindow();
                var helper = new System.Windows.Interop.WindowInteropHelper(win);
                helper.Owner = commandData.Application.MainWindowHandle;
                win.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var win = new AboutWindow();
                var helper = new System.Windows.Interop.WindowInteropHelper(win);
                helper.Owner = commandData.Application.MainWindowHandle;
                win.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public static class CommandHelper
    {
        public static bool EnsureActivated(UIApplication uiApp)
        {
            if (!LicenseManager.IsActivated())
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\QicTools\QicBoqMapper"))
                {
                    if (key != null)
                    {
                        object expiresAtVal = key.GetValue("ExpiresAt");
                        if (expiresAtVal != null && DateTimeOffset.TryParse(expiresAtVal.ToString(), out DateTimeOffset expiresAt))
                        {
                            if (DateTimeOffset.UtcNow > expiresAt)
                            {
                                TaskDialog.Show("License Expired", "License Expired. Please contact the administrator to renew your subscription.");
                            }
                        }
                    }
                }
            }

            if (LicenseManager.IsActivated())
                return true;

            var win = new LicenseWindow();
            var helper = new System.Windows.Interop.WindowInteropHelper(win);
            helper.Owner = uiApp.MainWindowHandle;
            bool? result = win.ShowDialog();

            return result == true && LicenseManager.IsActivated();
        }
    }
}
