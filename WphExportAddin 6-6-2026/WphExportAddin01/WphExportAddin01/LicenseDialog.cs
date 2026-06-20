using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WphExportAddin
{
    /// <summary>
    /// Simple WinForms dialog for WPH Export licensing.
    /// Collects user email and activation code to activate online.
    /// Includes a link to request activation codes directly from the owner.
    /// </summary>
    public class LicenseDialog : Form
    {
        private TextBox _emailBox;
        private TextBox _codeBox;
        private Button _activateBtn;
        private Button _cancelBtn;

        public LicenseDialog()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "WPH Export — License Activation";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ClientSize = new Size(380, 210);

            // Intro label
            var introLabel = new Label
            {
                Text = "Please enter your registered email and activation code to activate the WPH Export Add-in.",
                Left = 16, Top = 16, Width = 348, Height = 40,
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            Controls.Add(introLabel);

            int labelX = 16, fieldX = 130, y = 64, rowH = 34, fieldW = 234;

            // Email
            var emailLabel = new Label { Text = "Email:", Left = labelX, Top = y, Width = 110 };
            Controls.Add(emailLabel);
            _emailBox = new TextBox { Left = fieldX, Top = y - 3, Width = fieldW };
            Controls.Add(_emailBox);
            y += rowH;

            // Activation Code
            var codeLabel = new Label { Text = "Activation Code:", Left = labelX, Top = y, Width = 110 };
            Controls.Add(codeLabel);
            _codeBox = new TextBox { Left = fieldX, Top = y - 3, Width = fieldW };
            Controls.Add(_codeBox);
            y += rowH + 10;

            // Request Code Link
            var requestLink = new LinkLabel
            {
                Text = "Request Code",
                Left = labelX,
                Top = y + 5,
                Width = 100,
                Height = 20
            };
            requestLink.LinkClicked += (s, ev) =>
            {
                try
                {
                    string machineId = LicenseManager.GetMachineIdentifier();
                    string userEmail = _emailBox.Text.Trim();

                    // Format structured email message
                    var sb = new StringBuilder();
                    sb.AppendLine("Hello Ahmed,");
                    sb.AppendLine();
                    sb.AppendLine("I would like to request an activation code for the WPH Export Add-in.");
                    sb.AppendLine();
                    sb.AppendLine("--- Device Details ---");
                    sb.AppendLine("Machine ID: " + machineId);
                    if (!string.IsNullOrEmpty(userEmail))
                    {
                        sb.AppendLine("Email Address: " + userEmail);
                    }
                    else
                    {
                        sb.AppendLine("Email Address: (Not specified by user)");
                    }

                    string mailtoUrl = string.Format(
                        "mailto:engkhalaf7@gmail.com?subject=WPH%20Export%20Activation%20Request&body={0}",
                        Uri.EscapeDataString(sb.ToString())
                    );

                    System.Diagnostics.Process.Start(mailtoUrl);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open default email client:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            Controls.Add(requestLink);

            // Activate Button
            _activateBtn = new Button
            {
                Text = "Activate",
                Left = fieldX,
                Top = y,
                Width = 112,
                Height = 30,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _activateBtn.FlatAppearance.BorderSize = 0;
            _activateBtn.Click += OnActivate;
            Controls.Add(_activateBtn);

            // Cancel Button
            _cancelBtn = new Button
            {
                Text = "Cancel",
                Left = fieldX + 122,
                Top = y,
                Width = 112,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(_cancelBtn);

            AcceptButton = _activateBtn;
            CancelButton = _cancelBtn;
        }

        private void OnActivate(object sender, EventArgs e)
        {
            string email = _emailBox.Text.Trim();
            string code = _codeBox.Text.Trim();

            _activateBtn.Enabled = false;
            _cancelBtn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                if (LicenseManager.Activate(email, code, out string error))
                {
                    MessageBox.Show(
                        "Activation successful!\n\nThank you for licensing the WPH Export Add-in.",
                        "WPH Export — Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    MessageBox.Show(
                        "Activation failed:\n" + error,
                        "WPH Export — Activation Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            finally
            {
                _activateBtn.Enabled = true;
                _cancelBtn.Enabled = true;
                Cursor = Cursors.Default;
            }
        }
    }
}
