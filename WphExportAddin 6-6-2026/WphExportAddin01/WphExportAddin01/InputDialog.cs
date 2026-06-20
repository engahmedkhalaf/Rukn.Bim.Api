using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;

// ---------------------------------------------------------------------------
// Type aliases. Both Autodesk.Revit.DB and System.Windows.Forms/System.Drawing
// define types with the same simple names (Form, Color, View, ...). These
// aliases lock the ambiguous names to the UI versions for THIS FILE so we never
// have to fully-qualify them, and so the compiler can't pick the Revit type by
// mistake. (The Revit View type is still reachable as Autodesk.Revit.DB.View
// where the code needs it.)
// ---------------------------------------------------------------------------
using Form = System.Windows.Forms.Form;
using Color = System.Drawing.Color;

namespace WphExportAddin
{
    /// <summary>
    /// Minimal WinForms dialog to collect the export parameters at run time
    /// (the GUI equivalent of the Dynamo IN[] ports). Built entirely in code so
    /// there's no .Designer.cs / .resx to manage. On OK it writes the user's
    /// choices back into the ExportSettings instance passed to the constructor.
    /// </summary>
    public class InputDialog : Form
    {
        private readonly ExportSettings _settings;
        private readonly Document _doc;          // needed to list 3D views

        private ComboBox _modeBox;
        private TextBox _folderBox;
        private TextBox _disciplineBox;
        private ComboBox _ifcBox;
        private ComboBox _linksBox;              // links handling
        private NumericUpDown _delayBox;         // delay between each export
        private CheckedListBox _viewList;        // multi-select 3D view checkboxes
        private CheckBox _dateStampBox;
        private CheckBox _subfolderBox;

        public InputDialog(ExportSettings settings, Document doc)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            BuildUi();
            PopulateViews();
            LoadFromSettings();
        }

        private void BuildUi()
        {
            Text = "WPH Export";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(480, 510);

            int labelX = 16, fieldX = 150, y = 18, rowH = 34, fieldW = 200;

            // Run mode
            AddLabel("Run mode:", labelX, y);
            _modeBox = new ComboBox
            {
                Left = fieldX,
                Top = y - 3,
                Width = fieldW,
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
                {
                    if (fbd.ShowDialog() == DialogResult.OK)
                        _folderBox.Text = fbd.SelectedPath;
                }
            };
            Controls.Add(browse);
            y += rowH;

            // Model Name
            AddLabel("Model name:", labelX, y);
            _disciplineBox = new TextBox { Left = fieldX, Top = y - 3, Width = fieldW };
            Controls.Add(_disciplineBox);
            y += rowH;

            // IFC version
            AddLabel("IFC version:", labelX, y);
            _ifcBox = new ComboBox
            {
                Left = fieldX,
                Top = y - 3,
                Width = fieldW,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _ifcBox.Items.AddRange(ExportSettings.IfcVersionKeys.Cast<object>().ToArray());
            Controls.Add(_ifcBox);
            y += rowH;

            // Links handling
            AddLabel("Links:", labelX, y);
            _linksBox = new ComboBox
            {
                Left = fieldX,
                Top = y - 3,
                Width = fieldW,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _linksBox.Items.AddRange(new object[]
            {
                "None", "Include in host", "Separate files"
            });
            Controls.Add(_linksBox);
            y += rowH;

            // Delay between each export (applies between views/formats here).
            AddLabel("Delay between exports (s):", labelX, y);
            _delayBox = new NumericUpDown
            {
                Left = fieldX,
                Top = y - 3,
                Width = 70,
                Minimum = 0,
                Maximum = 600,
                Value = 0
            };
            Controls.Add(_delayBox);
            y += rowH;

            // 3D views (MULTI-SELECT). Tick one or more 3D views; both NWC and
            // IFC export one file per ticked view. Tick NONE = whole model.
            AddLabel("3D views:", labelX, y);
            _viewList = new CheckedListBox
            {
                Left = fieldX,
                Top = y - 3,
                Width = fieldW,
                Height = 110,
                CheckOnClick = true,                 // single click toggles the tick
                IntegralHeight = false
            };
            Controls.Add(_viewList);

            // Small "All" / "None" helpers next to the list.
            var allBtn = new Button { Text = "All", Left = fieldX + fieldW + 8, Top = y - 3, Width = 50 };
            allBtn.Click += (s, e) => SetAllChecked(true);
            Controls.Add(allBtn);
            var noneBtn = new Button { Text = "None", Left = fieldX + fieldW + 8, Top = y + 25, Width = 50 };
            noneBtn.Click += (s, e) => SetAllChecked(false);
            Controls.Add(noneBtn);

            // Hint under the list.
            var hint = new Label
            {
                Text = "(tick none = whole model)",
                Left = fieldX,
                Top = y + 110,
                Width = fieldW + 60,
                AutoSize = false,
                ForeColor = Color.Gray
            };
            Controls.Add(hint);
            y += 110 + 22;   // list height + hint

            // Date stamp
            _dateStampBox = new CheckBox
            {
                Text = "Append date stamp (-YYYY-MM-DD)",
                Left = fieldX,
                Top = y,
                Width = fieldW + 80
            };
            Controls.Add(_dateStampBox);
            y += rowH - 8;

            // Subfolders
            _subfolderBox = new CheckBox
            {
                Text = "Per-format subfolders (NWC/IFC)",
                Left = fieldX,
                Top = y,
                Width = fieldW + 80
            };
            Controls.Add(_subfolderBox);
            y += rowH;

            // OK / Cancel
            var ok = new Button
            {
                Text = "Export",
                Left = fieldX,
                Top = y,
                Width = 95,
                DialogResult = DialogResult.OK
            };
            ok.Click += OnOk;
            Controls.Add(ok);

            var cancel = new Button
            {
                Text = "Cancel",
                Left = fieldX + 105,
                Top = y,
                Width = 95,
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void AddLabel(string text, int x, int y)
        {
            Controls.Add(new Label
            {
                Text = text,
                Left = x,
                Top = y,
                Width = 130,
                AutoSize = false
            });
        }

        // 0=None, 1=Include in host, 2=Separate files
        internal static int LinkModeToIndex(LinkMode m)
        {
            switch (m)
            {
                case LinkMode.IncludeInHost: return 1;
                case LinkMode.SeparateFiles: return 2;
                default: return 0;
            }
        }

        internal static LinkMode IndexToLinkMode(int i)
        {
            switch (i)
            {
                case 1: return LinkMode.IncludeInHost;
                case 2: return LinkMode.SeparateFiles;
                default: return LinkMode.None;
            }
        }

        /// <summary>
        /// Fill the checkbox list with every real 3D view in the model (skipping
        /// view templates). Ticking none = whole model. Uses ViewChoice items.
        /// </summary>
        private void PopulateViews()
        {
            var views = new FilteredElementCollector(_doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => v != null && !v.IsTemplate)
                .OrderBy(v => v.Name)
                .Select(v => new ViewChoice { Name = v.Name, Id = v.Id })
                .ToList();

            foreach (var vc in views)
                _viewList.Items.Add(vc, false);   // added unchecked
        }

        /// <summary>Tick or untick every item in the view list.</summary>
        private void SetAllChecked(bool isChecked)
        {
            for (int i = 0; i < _viewList.Items.Count; i++)
                _viewList.SetItemChecked(i, isChecked);
        }

        private void LoadFromSettings()
        {
            _modeBox.SelectedItem = _settings.Mode == RunMode.Monthly ? "monthly" : "weekly";
            string docTitle = _doc.Title;
            if (docTitle.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                docTitle = docTitle.Substring(0, docTitle.Length - 4);
            _disciplineBox.Text = !string.IsNullOrWhiteSpace(_settings.ModelName) ? _settings.ModelName : docTitle;

            _ifcBox.SelectedItem = _ifcBox.Items.Contains(_settings.IfcVersionKey)
                ? (object)_settings.IfcVersionKey
                : (_ifcBox.Items.Count > 0 ? _ifcBox.Items[0] : null);

            _linksBox.SelectedIndex = LinkModeToIndex(_settings.Links);

            // Clamp the stored delay into the control's allowed range.
            int d = _settings.DelaySeconds;
            if (d < (int)_delayBox.Minimum) d = (int)_delayBox.Minimum;
            if (d > (int)_delayBox.Maximum) d = (int)_delayBox.Maximum;
            _delayBox.Value = d;

            _dateStampBox.Checked = _settings.DateStamp;
            _subfolderBox.Checked = _settings.UseSubfolders;
        }

        private void OnOk(object sender, EventArgs e)
        {
            // Basic validation; keep it light so the command stays runnable.
            if (string.IsNullOrWhiteSpace(_folderBox.Text))
            {
                MessageBox.Show("Please choose an output folder.", "WPH Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None; // keep dialog open
                return;
            }
            if (string.IsNullOrWhiteSpace(_disciplineBox.Text))
            {
                MessageBox.Show("Please enter a model name.", "WPH Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            // Write choices back into the shared settings object.
            _settings.Mode = string.Equals((string)_modeBox.SelectedItem, "monthly",
                StringComparison.OrdinalIgnoreCase) ? RunMode.Monthly : RunMode.Weekly;
            _settings.OutputFolder = _folderBox.Text.Trim();
            _settings.ModelName = _disciplineBox.Text.Trim();
            _settings.IfcVersionKey = _ifcBox.SelectedItem?.ToString() ?? "IFC2x3CV2";
            _settings.Links = IndexToLinkMode(_linksBox.SelectedIndex);
            _settings.DelaySeconds = (int)_delayBox.Value;
            _settings.DateStamp = _dateStampBox.Checked;
            _settings.UseSubfolders = _subfolderBox.Checked;

            // Capture every TICKED 3D view. Both NWC and IFC export one file per
            // view. If none are ticked, SelectedViews stays empty => whole model.
            _settings.SelectedViews = _viewList.CheckedItems
                .Cast<ViewChoice>()
                .ToList();
        }
    }
}
