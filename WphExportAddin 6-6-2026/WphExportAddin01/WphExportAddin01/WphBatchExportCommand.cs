using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Lock the ambiguous names to the WinForms/Drawing versions for THIS file,
// exactly like InputDialog.cs does (Revit also defines Form / Color / View).
using Form = System.Windows.Forms.Form;
using Color = System.Drawing.Color;
using ComboBox = System.Windows.Forms.ComboBox;
using TextBox = System.Windows.Forms.TextBox;

namespace WphExportAddin
{
    // =====================================================================
    // BATCH EXPORT
    // ---------------------------------------------------------------------
    // Lets the user pick MANY .rvt model files and export them one after the
    // other. Each model is opened in the background, handed to the SAME
    // WphExporter you already use (so NWC/IFC settings are byte-for-byte the
    // same), then closed. The delay from ExportSettings.DelaySeconds is applied
    // BETWEEN models (here) AND between views/formats (inside WphExporter), so
    // one setting paces every level of the loop.
    //
    // Error isolation is layered so the batch is "accurate" and never aborts:
    //   - per FORMAT  : handled inside WphExporter (NWC/IFC try-catch)
    //   - per MODEL   : handled in BatchExporter (open / export / close)
    // One bad model is logged and skipped; the rest still run.
    //
    // SCOPE NOTE: batch always exports the WHOLE MODEL (no per-view ticking),
    // because the 3D views differ from file to file and can't be picked up
    // front for files that aren't open yet.
    // =====================================================================

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WphBatchExportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
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

            // The batch settings + file list are collected in one dialog.
            var settings = new ExportSettings
            {
                // Batch is whole-model by definition (see SCOPE NOTE above).
                SelectedViews = new List<ViewChoice>()
            };

            List<string> files;

            using (var dlg = new BatchInputDialog(settings))
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                files = dlg.SelectedFiles;
                // dlg has written Mode/folder/discipline/IFC/DelaySeconds/etc.
                // back into `settings`.
            }

            if (files == null || files.Count == 0)
            {
                message = "No model files were selected.";
                return Result.Failed;
            }

            try
            {
                var batch = new BatchExporter(uiApp, settings, files);
                string log = batch.Run();

                new TaskDialog("WPH Export — Batch result")
                {
                    MainInstruction = "Batch export finished.",
                    MainContent = log,
                    CommonButtons = TaskDialogCommonButtons.Close
                }.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "WPH batch export failed: " + ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Code-only WinForms dialog for the batch run. Same hand-built style as
    /// InputDialog.cs (no .Designer.cs / .resx). Collects mode, folder,
    /// discipline, IFC version, the delay between exports, the on/off toggles,
    /// and the list of .rvt files to process.
    /// </summary>
    public class BatchInputDialog : Form
    {
        private readonly ExportSettings _settings;

        private ComboBox _modeBox;
        private TextBox _folderBox;
        private TextBox _disciplineBox;
        private ComboBox _ifcBox;
        private NumericUpDown _delayBox;
        private CheckBox _dateStampBox;
        private CheckBox _subfolderBox;
        private ListBox _fileList;

        /// <summary>Full paths of the .rvt files to export (read after OK).</summary>
        public List<string> SelectedFiles { get; private set; } = new List<string>();

        public BatchInputDialog(ExportSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            BuildUi();
            LoadFromSettings();
        }

        private void BuildUi()
        {
            Text = "WPH Batch Export";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 470);

            int labelX = 16, fieldX = 150, y = 18, rowH = 34, fieldW = 240;

            // Run mode
            AddLabel("Run mode:", labelX, y);
            _modeBox = new ComboBox
            {
                Left = fieldX, Top = y - 3, Width = fieldW,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _modeBox.Items.AddRange(new object[] { "weekly", "monthly" });
            Controls.Add(_modeBox);
            y += rowH;

            // Output folder + browse
            AddLabel("Output folder:", labelX, y);
            _folderBox = new TextBox { Left = fieldX, Top = y - 3, Width = fieldW - 40 };
            Controls.Add(_folderBox);
            var browse = new Button { Text = "…", Left = fieldX + fieldW - 34, Top = y - 4, Width = 34 };
            browse.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog())
                    if (fbd.ShowDialog() == DialogResult.OK) _folderBox.Text = fbd.SelectedPath;
            };
            Controls.Add(browse);
            y += rowH;

            // Model Name
            AddLabel("Model name override:", labelX, y);
            _disciplineBox = new TextBox { Left = fieldX, Top = y - 3, Width = fieldW };
            Controls.Add(_disciplineBox);
            y += rowH;

            // IFC version
            AddLabel("IFC version:", labelX, y);
            _ifcBox = new ComboBox
            {
                Left = fieldX, Top = y - 3, Width = fieldW,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _ifcBox.Items.AddRange(ExportSettings.IfcVersionKeys.Cast<object>().ToArray());
            Controls.Add(_ifcBox);
            y += rowH;

            // Delay between exports
            AddLabel("Delay between exports (s):", labelX, y);
            _delayBox = new NumericUpDown
            {
                Left = fieldX, Top = y - 3, Width = 80,
                Minimum = 0, Maximum = 600, Value = 5
            };
            Controls.Add(_delayBox);
            var delayHint = new Label
            {
                Text = "(pause so Revit can settle)",
                Left = fieldX + 90, Top = y, Width = fieldW, AutoSize = false,
                ForeColor = Color.Gray
            };
            Controls.Add(delayHint);
            y += rowH;

            // Toggles
            _dateStampBox = new CheckBox
            {
                Text = "Append date stamp (-YYYY-MM-DD)",
                Left = fieldX, Top = y, Width = fieldW + 80
            };
            Controls.Add(_dateStampBox);
            y += rowH - 8;

            _subfolderBox = new CheckBox
            {
                Text = "Per-format subfolders (NWC/IFC)",
                Left = fieldX, Top = y, Width = fieldW + 80
            };
            Controls.Add(_subfolderBox);
            y += rowH;

            // File list + add/remove
            AddLabel("Models (.rvt):", labelX, y);
            _fileList = new ListBox
            {
                Left = fieldX, Top = y - 3, Width = fieldW, Height = 96,
                HorizontalScrollbar = true, SelectionMode = SelectionMode.MultiExtended
            };
            Controls.Add(_fileList);

            var addBtn = new Button { Text = "Add…", Left = fieldX + fieldW + 8, Top = y - 3, Width = 60 };
            addBtn.Click += (s, e) => AddFiles();
            Controls.Add(addBtn);

            var removeBtn = new Button { Text = "Remove", Left = fieldX + fieldW + 8, Top = y + 27, Width = 60 };
            removeBtn.Click += (s, e) => RemoveSelected();
            Controls.Add(removeBtn);

            var clearBtn = new Button { Text = "Clear", Left = fieldX + fieldW + 8, Top = y + 57, Width = 60 };
            clearBtn.Click += (s, e) => _fileList.Items.Clear();
            Controls.Add(clearBtn);

            y += 96 + 12;

            // OK / Cancel
            var ok = new Button
            {
                Text = "Export all", Left = fieldX, Top = y, Width = 100,
                DialogResult = DialogResult.OK
            };
            ok.Click += OnOk;
            Controls.Add(ok);

            var cancel = new Button
            {
                Text = "Cancel", Left = fieldX + 110, Top = y, Width = 95,
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void AddLabel(string text, int x, int y)
        {
            Controls.Add(new Label { Text = text, Left = x, Top = y, Width = 132, AutoSize = false });
        }

        private void AddFiles()
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "Select Revit models to export",
                Filter = "Revit models (*.rvt)|*.rvt",
                Multiselect = true
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                foreach (var f in ofd.FileNames)
                {
                    // Avoid duplicates in the queue.
                    if (!_fileList.Items.Cast<string>()
                            .Any(x => string.Equals(x, f, StringComparison.OrdinalIgnoreCase)))
                        _fileList.Items.Add(f);
                }
            }
        }

        private void RemoveSelected()
        {
            // Remove from the end so indices stay valid while deleting.
            var idx = _fileList.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
            foreach (int i in idx) _fileList.Items.RemoveAt(i);
        }

        private void LoadFromSettings()
        {
            _modeBox.SelectedItem = _settings.Mode == RunMode.Monthly ? "monthly" : "weekly";
            _folderBox.Text = _settings.OutputFolder;
            _disciplineBox.Text = !string.IsNullOrWhiteSpace(_settings.ModelName) ? _settings.ModelName : "<Use Model Name>";
            _ifcBox.SelectedItem = _ifcBox.Items.Contains(_settings.IfcVersionKey)
                ? (object)_settings.IfcVersionKey
                : (_ifcBox.Items.Count > 0 ? _ifcBox.Items[0] : null);
            _dateStampBox.Checked = _settings.DateStamp;
            _subfolderBox.Checked = _settings.UseSubfolders;
        }

        private void OnOk(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_folderBox.Text))
            {
                Warn("Please choose an output folder.");
                DialogResult = DialogResult.None; return;
            }
            if (string.IsNullOrWhiteSpace(_disciplineBox.Text))
            {
                Warn("Please enter a model name or use '<Use Model Name>'.");
                DialogResult = DialogResult.None; return;
            }
            if (_fileList.Items.Count == 0)
            {
                Warn("Add at least one .rvt model to the list.");
                DialogResult = DialogResult.None; return;
            }

            _settings.Mode = string.Equals((string)_modeBox.SelectedItem, "monthly",
                StringComparison.OrdinalIgnoreCase) ? RunMode.Monthly : RunMode.Weekly;
            _settings.OutputFolder = _folderBox.Text.Trim();
            _settings.ModelName = _disciplineBox.Text.Trim();
            _settings.IfcVersionKey = _ifcBox.SelectedItem?.ToString() ?? "IFC2x3CV2";
            _settings.DelaySeconds = (int)_delayBox.Value;
            _settings.DateStamp = _dateStampBox.Checked;
            _settings.UseSubfolders = _subfolderBox.Checked;

            SelectedFiles = _fileList.Items.Cast<string>().ToList();
        }

        private static void Warn(string text)
        {
            MessageBox.Show(text, "WPH Batch Export",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
