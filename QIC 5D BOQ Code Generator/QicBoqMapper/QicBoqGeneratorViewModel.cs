using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QicBoqMapper
{
    public class QicBoqGeneratorViewModel : INotifyPropertyChanged
    {
        private Document _doc;
        private UIDocument? _uiDoc;
        private readonly ExternalEvent _externalEvent;
        private readonly QicBoqExternalEventHandler _eventHandler;

        private string _excelFilePath = string.Empty;
        private string _selectionMode = "Entire Model";
        private string _matchingMethod = "Category + Family + Type";
        private string _separatorStyle = "Dash";
        private bool _caseInsensitive = true;
        private string _statusText = "Ready";
        private int _progressValue = 0;

        private bool _isCurrentSelectionSelected;
        private bool _isActiveViewSelected = true;
        private bool _isWholeModelSelected;

        private string _licenseKey = string.Empty;
        private string _licenseStatus = "License: Inactive";

        private string _categoryFilterText = string.Empty;
        private string _familyFilterText = string.Empty;
        private string _typeFilterText = string.Empty;

        private int _loadedRecordsCount = 0;
        private int _matchedTypesCount = 0;
        private int _unmatchedTypesCount = 0;
        private int _processedElementsCount = 0;
        private int _warningCount = 0;
        private int _errorCount = 0;

        private List<BoqRecord> _excelRecords = new List<BoqRecord>();
        private List<AuditRecord> _auditRecords = new List<AuditRecord>();

        public QicBoqGeneratorViewModel(Document doc, UIDocument? uiDoc, ExternalEvent externalEvent, QicBoqExternalEventHandler eventHandler)
        {
            _doc = doc;
            _uiDoc = uiDoc;
            _externalEvent = externalEvent;
            _eventHandler = eventHandler;

            SelectionModes = new ObservableCollection<string> { "Current Selection", "Active View", "Entire Model" };
            MatchingMethods = new ObservableCollection<string> { "Category + Family + Type", "Element ID" };
            SeparatorStyles = new ObservableCollection<string> { "- (Dash)", ". (Dot)", "_ (Underscore)", ", (Comma)" };

            // Load persisted settings and Excel data
            _excelFilePath = QicBoqApp.LastExcelFilePath;
            _excelRecords = QicBoqApp.LastExcelRecords;
            _selectionMode = QicBoqApp.LastSelectionMode;
            _matchingMethod = QicBoqApp.LastMatchingMethod;
            _separatorStyle = QicBoqApp.LastSeparatorStyle;
            if (_separatorStyle == "Dash" || string.IsNullOrEmpty(_separatorStyle)) _separatorStyle = "- (Dash)";
            else if (_separatorStyle == "Dot") _separatorStyle = ". (Dot)";
            else if (_separatorStyle == "Underscore") _separatorStyle = "_ (Underscore)";
            else if (_separatorStyle == "Comma") _separatorStyle = ", (Comma)";
            _caseInsensitive = QicBoqApp.LastCaseInsensitive;

            if (_selectionMode == "Current Selection")
            {
                _isCurrentSelectionSelected = true;
                _isActiveViewSelected = false;
                _isWholeModelSelected = false;
            }
            else if (_selectionMode == "Entire Model" || _selectionMode == "Whole Model")
            {
                _isCurrentSelectionSelected = false;
                _isActiveViewSelected = false;
                _isWholeModelSelected = true;
            }
            else
            {
                _isCurrentSelectionSelected = false;
                _isActiveViewSelected = true;
                _isWholeModelSelected = false;
            }

            if (_excelRecords.Count > 0)
            {
                _loadedRecordsCount = _excelRecords.Count;
                _statusText = $"Ready (Loaded {LoadedRecordsCount} BOQ records)";
            }

            CategoryItems = new ObservableCollection<CategorySelectionItem>();
            foreach (var name in CategoryMappingService.GetCategoryNames())
            {
                var item = new CategorySelectionItem { Name = name, IsSelected = true };
                item.PropertyChanged += (s, e) => {
                    if (e.PropertyName == nameof(CategorySelectionItem.IsSelected))
                    {
                        OnPropertyChanged(nameof(SelectedCategoriesSummary));
                    }
                };
                CategoryItems.Add(item);
            }

            BrowseFileCommand = new RelayCommand(BrowseFile);
            LoadExcelCommand = new RelayCommand(LoadExcel, CanLoadExcel);
            ImportExcelCommand = new RelayCommand(ImportExcel);
            ExportElementsCommand = new RelayCommand(ExportElements, CanExportElements);
            ValidateCommand = new RelayCommand(ValidateMapping, CanValidateOrGenerate);
            GenerateCommand = new RelayCommand(GenerateCodes, CanValidateOrGenerate);
            ExportReportCommand = new RelayCommand(ExportReport, CanExportReport);
            ActivateLicenseCommand = new RelayCommand(ActivateLicense);
        }

        public void UpdateActiveDocument(Document doc, UIDocument? uiDoc)
        {
            _doc = doc;
            _uiDoc = uiDoc;
        }

        public void TriggerAction(string action)
        {
            switch (action.ToLower())
            {
                case "export":
                    if (ExportElementsCommand.CanExecute(null))
                        ExportElementsCommand.Execute(null);
                    break;
                case "import":
                    if (BrowseFileCommand.CanExecute(null))
                        BrowseFileCommand.Execute(null);
                    break;
                case "validate":
                    if (ValidateCommand.CanExecute(null))
                        ValidateCommand.Execute(null);
                    else
                        StatusText = "Please import a mapping Excel file first to validate.";
                    break;
                case "generate":
                    if (GenerateCommand.CanExecute(null))
                        GenerateCommand.Execute(null);
                    else
                        StatusText = "Please import a mapping Excel file first to generate codes.";
                    break;
                case "audit":
                    if (ExportReportCommand.CanExecute(null))
                        ExportReportCommand.Execute(null);
                    else
                        StatusText = "No audit records to export. Please validate or generate first.";
                    break;
            }
        }

        // Selection Scope Radio Button Properties
        public bool IsCurrentSelectionSelected
        {
            get => _isCurrentSelectionSelected;
            set
            {
                if (_isCurrentSelectionSelected != value)
                {
                    _isCurrentSelectionSelected = value;
                    if (value)
                    {
                        IsActiveViewSelected = false;
                        IsWholeModelSelected = false;
                        SelectionMode = "Current Selection";
                    }
                    OnPropertyChanged(nameof(IsCurrentSelectionSelected));
                }
            }
        }

        public bool IsActiveViewSelected
        {
            get => _isActiveViewSelected;
            set
            {
                if (_isActiveViewSelected != value)
                {
                    _isActiveViewSelected = value;
                    if (value)
                    {
                        IsCurrentSelectionSelected = false;
                        IsWholeModelSelected = false;
                        SelectionMode = "Active View";
                    }
                    OnPropertyChanged(nameof(IsActiveViewSelected));
                }
            }
        }

        public bool IsWholeModelSelected
        {
            get => _isWholeModelSelected;
            set
            {
                if (_isWholeModelSelected != value)
                {
                    _isWholeModelSelected = value;
                    if (value)
                    {
                        IsCurrentSelectionSelected = false;
                        IsActiveViewSelected = false;
                        SelectionMode = "Entire Model";
                    }
                    OnPropertyChanged(nameof(IsWholeModelSelected));
                }
            }
        }

        // License Properties
        public string LicenseKey
        {
            get => _licenseKey;
            set { _licenseKey = value; OnPropertyChanged(nameof(LicenseKey)); }
        }
        
        public string LicenseStatus
        {
            get => _licenseStatus;
            set { _licenseStatus = value; OnPropertyChanged(nameof(LicenseStatus)); }
        }

        // Selection Filters
        public string CategoryFilterText
        {
            get => _categoryFilterText;
            set { _categoryFilterText = value; OnPropertyChanged(nameof(CategoryFilterText)); }
        }

        public string FamilyFilterText
        {
            get => _familyFilterText;
            set { _familyFilterText = value; OnPropertyChanged(nameof(FamilyFilterText)); }
        }

        public string TypeFilterText
        {
            get => _typeFilterText;
            set { _typeFilterText = value; OnPropertyChanged(nameof(TypeFilterText)); }
        }

        public string SelectedCategoriesSummary
        {
            get
            {
                var selected = CategoryItems.Where(c => c.IsSelected).Select(c => c.Name).ToList();
                if (selected.Count == 0)
                {
                    return "None Selected";
                }
                if (selected.Count == CategoryItems.Count)
                {
                    return "All Categories";
                }
                if (selected.Count == 1)
                {
                    return selected[0];
                }
                return $"{selected.Count} Categories Selected";
            }
        }

        // Properties
        public string ExcelFilePath
        {
            get => _excelFilePath;
            set { _excelFilePath = value; QicBoqApp.LastExcelFilePath = value; OnPropertyChanged(nameof(ExcelFilePath)); CommandManager.InvalidateRequerySuggested(); }
        }

        public ObservableCollection<string> SelectionModes { get; }
        public string SelectionMode
        {
            get => _selectionMode;
            set { _selectionMode = value; QicBoqApp.LastSelectionMode = value; OnPropertyChanged(nameof(SelectionMode)); }
        }

        public ObservableCollection<string> MatchingMethods { get; }
        public string MatchingMethod
        {
            get => _matchingMethod;
            set { _matchingMethod = value; QicBoqApp.LastMatchingMethod = value; OnPropertyChanged(nameof(MatchingMethod)); }
        }

        public ObservableCollection<CategorySelectionItem> CategoryItems { get; }

        public ObservableCollection<string> SeparatorStyles { get; }
        public string SeparatorStyle
        {
            get => _separatorStyle;
            set { _separatorStyle = value; QicBoqApp.LastSeparatorStyle = value; OnPropertyChanged(nameof(SeparatorStyle)); }
        }

        public bool CaseInsensitive
        {
            get => _caseInsensitive;
            set { _caseInsensitive = value; QicBoqApp.LastCaseInsensitive = value; OnPropertyChanged(nameof(CaseInsensitive)); }
        }

        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(nameof(StatusText)); }
        }

        public int ProgressValue
        {
            get => _progressValue;
            set { _progressValue = value; OnPropertyChanged(nameof(ProgressValue)); }
        }

        public int LoadedRecordsCount
        {
            get => _loadedRecordsCount;
            set { _loadedRecordsCount = value; OnPropertyChanged(nameof(LoadedRecordsCount)); }
        }

        public int MatchedTypesCount
        {
            get => _matchedTypesCount;
            set { _matchedTypesCount = value; OnPropertyChanged(nameof(MatchedTypesCount)); }
        }

        public int UnmatchedTypesCount
        {
            get => _unmatchedTypesCount;
            set { _unmatchedTypesCount = value; OnPropertyChanged(nameof(UnmatchedTypesCount)); }
        }

        public int ProcessedElementsCount
        {
            get => _processedElementsCount;
            set { _processedElementsCount = value; OnPropertyChanged(nameof(ProcessedElementsCount)); }
        }

        public int WarningCount
        {
            get => _warningCount;
            set { _warningCount = value; OnPropertyChanged(nameof(WarningCount)); }
        }

        public int ErrorCount
        {
            get => _errorCount;
            set { _errorCount = value; OnPropertyChanged(nameof(ErrorCount)); }
        }

        // Commands
        public ICommand BrowseFileCommand { get; }
        public ICommand LoadExcelCommand { get; }
        public ICommand ImportExcelCommand { get; }
        public ICommand ExportElementsCommand { get; }
        public ICommand ValidateCommand { get; }
        public ICommand GenerateCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand ActivateLicenseCommand { get; }

        private void BrowseFile(object parameter)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*",
                Title = "Select Excel BOQ Mapping File"
            };

            if (dlg.ShowDialog() == true)
            {
                ExcelFilePath = dlg.FileName;
            }
        }

        private void ImportExcel(object parameter)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*",
                Title = "Select Excel BOQ Mapping File"
            };

            if (dlg.ShowDialog() == true)
            {
                ExcelFilePath = dlg.FileName;
                LoadExcel(null);
            }
        }

        private bool CanLoadExcel(object parameter)
        {
            return !string.IsNullOrWhiteSpace(ExcelFilePath) && File.Exists(ExcelFilePath);
        }

        private void LoadExcel(object parameter)
        {
            try
            {
                ProgressValue = 10;
                StatusText = "Reading Excel file...";
                _excelRecords = QicBoqExcelService.LoadBoqRecords(ExcelFilePath);
                LoadedRecordsCount = _excelRecords.Count;

                QicBoqApp.LastExcelFilePath = ExcelFilePath;
                QicBoqApp.LastExcelRecords = _excelRecords;

                ProgressValue = 100;
                StatusText = $"Loaded {LoadedRecordsCount} BOQ records successfully.";
            }
            catch (Exception ex)
            {
                ProgressValue = 0;
                StatusText = "Excel Load Error.";
                MessageBox.Show($"Failed to load Excel file:\n{ex.Message}", "Excel Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanExportElements(object parameter)
        {
            return CategoryItems.Any(c => c.IsSelected);
        }

        private void ExportElements(object parameter)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                Title = "Save Exported Elements As",
                FileName = "QIC_BOQ_Exported_Elements"
            };

            if (sfd.ShowDialog() == true)
            {
                _eventHandler.QueueAction(uiApp =>
                {
                    try
                    {
                        ProgressValue = 20;
                        StatusText = "Collecting Revit elements...";
                        var selectedCats = CategoryItems.Where(c => c.IsSelected).Select(c => c.Name).ToList();
                        var elements = QicBoqManager.GetElements(_doc, _uiDoc, SelectionMode, selectedCats, CategoryFilterText, FamilyFilterText, TypeFilterText);

                        ProgressValue = 50;
                        StatusText = "Exporting to Excel...";
                        QicBoqExportService.ExportElements(sfd.FileName, _doc, elements);

                        ProgressValue = 100;
                        StatusText = "Elements exported successfully.";
                        MessageBox.Show($"Successfully exported {elements.Count} elements to Excel!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        ProgressValue = 0;
                        StatusText = "Export Error.";
                        MessageBox.Show($"Failed to export elements:\n{ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                });
                _externalEvent.Raise();
            }
        }

        private bool CanValidateOrGenerate(object parameter)
        {
            return LicenseManager.IsActivated() && _excelRecords.Count > 0 && CategoryItems.Any(c => c.IsSelected);
        }

        private void ValidateMapping(object parameter)
        {
            _eventHandler.QueueAction(uiApp =>
            {
                try
                {
                    ProgressValue = 20;
                    StatusText = "Collecting Revit elements...";
                    var selectedCats = CategoryItems.Where(c => c.IsSelected).Select(c => c.Name).ToList();
                    var elements = QicBoqManager.GetElements(_doc, _uiDoc, SelectionMode, selectedCats, CategoryFilterText, FamilyFilterText, TypeFilterText);

                    ProgressValue = 50;
                    StatusText = $"Mapping {elements.Count} elements...";
                    int matched, unmatched;
                    _auditRecords = QicBoqManager.PerformMapping(_doc, elements, _excelRecords, SeparatorStyle, MatchingMethod, CaseInsensitive, out matched, out unmatched);

                    MatchedTypesCount = matched;
                    UnmatchedTypesCount = unmatched;
                    ProcessedElementsCount = elements.Count;
                    WarningCount = _auditRecords.Count(r => r.Status == "Warning");
                    ErrorCount = _auditRecords.Count(r => r.Status == "Error");

                    ProgressValue = 100;
                    StatusText = $"Validation completed. Matched: {matched}, Errors: {ErrorCount}.";
                }
                catch (Exception ex)
                {
                    ProgressValue = 0;
                    StatusText = "Validation Error.";
                    MessageBox.Show($"Failed to validate mapping:\n{ex.Message}", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
            _externalEvent.Raise();
        }

        private void GenerateCodes(object parameter)
        {
            if (!LicenseManager.IsActivated())
            {
                var win = new LicenseWindow();
                var activeWin = System.Windows.Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive)
                                ?? System.Windows.Application.Current.MainWindow;
                if (activeWin != null)
                {
                    win.Owner = activeWin;
                }

                bool? result = win.ShowDialog();
                if (result != true || !LicenseManager.IsActivated())
                {
                    StatusText = "Activation required to generate 5D BOQ codes.";
                    return;
                }
            }

            _eventHandler.QueueAction(uiApp =>
            {
                try
                {
                    ProgressValue = 10;
                    StatusText = "Creating shared parameters...";
                    var selectedCats = CategoryItems.Where(c => c.IsSelected).Select(c => c.Name).ToList();
                    QicBoqManager.CreateSharedParameters(_doc, selectedCats);

                    ProgressValue = 30;
                    StatusText = "Gathering and mapping elements...";
                    var elements = QicBoqManager.GetElements(_doc, _uiDoc, SelectionMode, selectedCats, CategoryFilterText, FamilyFilterText, TypeFilterText);
                    int matched, unmatched;
                    _auditRecords = QicBoqManager.PerformMapping(_doc, elements, _excelRecords, SeparatorStyle, MatchingMethod, CaseInsensitive, out matched, out unmatched);

                    ProgressValue = 60;
                    StatusText = "Updating parameters...";
                    QicBoqManager.UpdateParameters(_doc, elements, _auditRecords, _excelRecords, MatchingMethod, CaseInsensitive);

                    MatchedTypesCount = matched;
                    UnmatchedTypesCount = unmatched;
                    ProcessedElementsCount = elements.Count;
                    WarningCount = _auditRecords.Count(r => r.Status == "Warning");
                    ErrorCount = _auditRecords.Count(r => r.Status == "Error");

                    ProgressValue = 100;
                    StatusText = $"BOQ Codes generated successfully for {elements.Count} elements.";
                    MessageBox.Show(
                        $"Elements Processed: {elements.Count}\n" +
                        $"Elements Updated: {MatchedTypesCount}\n" +
                        $"Elements Skipped: {ErrorCount}",
                        "Update Summary",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    ProgressValue = 0;
                    StatusText = "Execution Error.";
                    MessageBox.Show($"Failed to generate codes:\n{ex.Message}", "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
            _externalEvent.Raise();
        }

        private bool CanExportReport(object parameter)
        {
            return _auditRecords.Count > 0;
        }

        private void ExportReport(object parameter)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV UTF-8 (*.csv)|*.csv|JSON File (*.json)|*.json",
                Title = "Save Audit Report As",
                FileName = "QIC_BOQ_Audit_Report"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    string ext = Path.GetExtension(sfd.FileName).ToLower();
                    if (ext == ".xlsx") QicBoqAuditExporter.ExportToExcel(sfd.FileName, _auditRecords);
                    else if (ext == ".csv") QicBoqAuditExporter.ExportToCsv(sfd.FileName, _auditRecords);
                    else if (ext == ".json") QicBoqAuditExporter.ExportToJson(sfd.FileName, _auditRecords);

                    StatusText = "Audit report exported successfully.";
                    MessageBox.Show("Audit report saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to export report:\n{ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ActivateLicense(object parameter)
        {
            if (string.IsNullOrWhiteSpace(LicenseKey))
            {
                MessageBox.Show("Please enter a license key.", "License Activation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Simple mock activation validation. E.g. key length >= 5
            if (LicenseKey.Trim().Length >= 5)
            {
                LicenseStatus = "License: Active";
                MessageBox.Show("License activated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Invalid license key. Please try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class CategorySelectionItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public string Name { get; set; } = string.Empty;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;

        public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);
        public void Execute(object parameter) => _execute(parameter);
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
