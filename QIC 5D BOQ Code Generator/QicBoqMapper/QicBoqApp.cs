using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace QicBoqMapper
{
    public class QicBoqApp : IExternalApplication
    {
        private const string TabName = "QIC Tools";
        private const string PanelName = "5D BOQ Management";

        private static QicBoqGeneratorWindow? _mainWindow = null;
        private static QicBoqExternalEventHandler? _eventHandler = null;
        private static ExternalEvent? _externalEvent = null;

        public static string LastExcelFilePath { get; set; } = string.Empty;
        public static List<BoqRecord> LastExcelRecords { get; set; } = new List<BoqRecord>();
        public static string LastSelectionMode { get; set; } = "Entire Model";
        public static string LastMatchingMethod { get; set; } = "Category + Family + Type";
        public static string LastSeparatorStyle { get; set; } = "Dash";
        public static bool LastCaseInsensitive { get; set; } = true;

        public Result OnStartup(UIControlledApplication application)
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            application.Idling += OnIdling;
            try
            {
                _eventHandler = new QicBoqExternalEventHandler();
                _externalEvent = ExternalEvent.Create(_eventHandler);

                try { application.CreateRibbonTab(TabName); } catch { }

                var panel = application.CreateRibbonPanel(TabName, PanelName);
                string asmPath = Assembly.GetExecutingAssembly().Location;

                panel.AddItem(CreateButton(asmPath, "QicBoqManager", "QIC 5D BOQ\nManager", "QicBoqMapper.QicBoqManagerCommand", "Open the QIC 5D BOQ Manager workflow window."));
                panel.AddItem(CreateButton(asmPath, "License", "License", "QicBoqMapper.LicenseCommand", "Activate your product license."));
                panel.AddItem(CreateButton(asmPath, "About", "About", "QicBoqMapper.AboutCommand", "View add-in version and support details."));

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("QIC Tools Startup Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            AppDomain.CurrentDomain.AssemblyResolve -= OnAssemblyResolve;
            application.Idling -= OnIdling;

            if (_mainWindow != null && _mainWindow.IsVisible)
            {
                _mainWindow.Close();
            }
            return Result.Succeeded;
        }

        private static void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            if (_mainWindow != null && _mainWindow.IsVisible)
            {
                var uiApp = sender as UIApplication;
                var activeUiDoc = uiApp?.ActiveUIDocument;
                if (activeUiDoc != null)
                {
                    var vm = _mainWindow.DataContext as QicBoqGeneratorViewModel;
                    vm?.UpdateActiveDocument(activeUiDoc.Document, activeUiDoc);
                }
            }
        }

        public static void ShowWindow(UIApplication uiApp, string? action = null)
        {
            try
            {
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc?.Document;
                if (doc == null)
                {
                    TaskDialog.Show("QIC Tools", "No active document.");
                    return;
                }

                if (_mainWindow != null && _mainWindow.IsLoaded)
                {
                    var vm = _mainWindow.DataContext as QicBoqGeneratorViewModel;
                    if (vm != null)
                    {
                        vm.UpdateActiveDocument(doc, uiDoc);
                        if (!string.IsNullOrEmpty(action))
                        {
                            vm.TriggerAction(action!);
                        }
                    }
                    _mainWindow.Activate();
                    if (_mainWindow.WindowState == System.Windows.WindowState.Minimized)
                    {
                        _mainWindow.WindowState = System.Windows.WindowState.Normal;
                    }
                    return;
                }

                if (_externalEvent == null || _eventHandler == null)
                {
                    _eventHandler = new QicBoqExternalEventHandler();
                    _externalEvent = ExternalEvent.Create(_eventHandler);
                }

                var viewModel = new QicBoqGeneratorViewModel(doc, uiDoc, _externalEvent, _eventHandler);
                _mainWindow = new QicBoqGeneratorWindow(viewModel);

                var helper = new System.Windows.Interop.WindowInteropHelper(_mainWindow);
                helper.Owner = uiApp.MainWindowHandle;

                _mainWindow.Closed += (s, e) => _mainWindow = null;
                _mainWindow.Show();

                if (!string.IsNullOrEmpty(action))
                {
                    viewModel.TriggerAction(action!);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("QIC Tools Error", $"Failed to open window:\n{ex.Message}");
            }
        }

        private static PushButtonData CreateButton(string asmPath, string name, string text, string className, string tooltip)
        {
            var btn = new PushButtonData(name, text, asmPath, className)
            {
                ToolTip = tooltip
            };
            AttachIcons(btn, name);
            return btn;
        }

        private static void AttachIcons(PushButtonData btn, string baseName)
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            string icon32 = Path.Combine(dir, "Resources", baseName + "_32.png");
            string icon16 = Path.Combine(dir, "Resources", baseName + "_16.png");

            // Fallback icons if specific files don't exist
            if (!File.Exists(icon32)) icon32 = Path.Combine(dir, "Resources", "QicBoqManager_32.png");
            if (!File.Exists(icon16)) icon16 = Path.Combine(dir, "Resources", "QicBoqManager_16.png");

            if (File.Exists(icon32))
                btn.LargeImage = new BitmapImage(new Uri(icon32, UriKind.Absolute));
            if (File.Exists(icon16))
                btn.Image = new BitmapImage(new Uri(icon16, UriKind.Absolute));
        }

        private static Assembly? OnAssemblyResolve(object? sender, ResolveEventArgs args)
        {
            string assemblyName = new AssemblyName(args.Name).Name;
            if (assemblyName.Equals("EPPlus", StringComparison.OrdinalIgnoreCase) ||
                assemblyName.StartsWith("Microsoft.IO.RecyclableMemoryStream", StringComparison.OrdinalIgnoreCase))
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
                string dllPath = Path.Combine(dir, assemblyName + ".dll");
                if (File.Exists(dllPath))
                {
                    try
                    {
                        byte[] assemblyBytes = File.ReadAllBytes(dllPath);
                        return Assembly.Load(assemblyBytes);
                    }
                    catch
                    {
                        return Assembly.LoadFrom(dllPath);
                    }
                }
            }
            return null;
        }
    }
}
