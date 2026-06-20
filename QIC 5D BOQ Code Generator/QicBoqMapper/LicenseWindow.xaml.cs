using System;
using System.Windows;
using System.Threading.Tasks;

namespace QicBoqMapper
{
    public partial class LicenseWindow : Window
    {
        public LicenseWindow()
        {
            InitializeComponent();
            
            if (LicenseManager.IsActivated())
            {
                StatusLabel.Text = "Status: Activated successfully.";
                StatusLabel.Foreground = System.Windows.Media.Brushes.LightGreen;
                StatusLabel.Visibility = Visibility.Visible;
            }
        }

        private async void ActivateButton_Click(object sender, RoutedEventArgs e)
        {
            string email = EmailTextBox.Text;
            string code = CodeTextBox.Text;

            if (!LicenseManager.ValidateInput(email, code))
            {
                StatusLabel.Text = "Error: Invalid email address or activation code (min 8 characters).";
                StatusLabel.Foreground = System.Windows.Media.Brushes.LightPink;
                StatusLabel.Visibility = Visibility.Visible;
                MessageBox.Show("Validation failed. Please verify your email and code.", "License Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                this.IsEnabled = false;
                StatusLabel.Text = "Validating with Supabase... Please wait.";
                StatusLabel.Foreground = System.Windows.Media.Brushes.Orange;
                StatusLabel.Visibility = Visibility.Visible;

                var validationResult = await Task.Run(() => LicenseManager.ValidateLicenseWithSupabaseAsync(email, code));
                bool isValid = validationResult.Item1;
                string expiresAtStr = validationResult.Item2;

                this.IsEnabled = true;

                if (isValid)
                {
                    LicenseManager.SaveLicense(email, code, true, expiresAtStr);
                    StatusLabel.Text = "Status: Activated successfully.";
                    StatusLabel.Foreground = System.Windows.Media.Brushes.LightGreen;
                    StatusLabel.Visibility = Visibility.Visible;
                    MessageBox.Show("Product license activated successfully!", "License Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.DialogResult = true;
                    this.Close();
                }
                else
                {
                    LicenseManager.SaveLicense(email, code, false);
                    StatusLabel.Text = "Error: License not found, inactive, or expired.";
                    StatusLabel.Foreground = System.Windows.Media.Brushes.LightPink;
                    StatusLabel.Visibility = Visibility.Visible;
                    MessageBox.Show("Invalid, inactive, or expired license. Verification failed.", "License Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                this.IsEnabled = true;
                try { LicenseManager.SaveLicense(email, code, false); } catch { }
                StatusLabel.Text = $"Error: {ex.Message}";
                StatusLabel.Foreground = System.Windows.Media.Brushes.LightPink;
                StatusLabel.Visibility = Visibility.Visible;
                MessageBox.Show($"An error occurred during license activation:\n{ex.Message}", "License Manager", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void RequestCode_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string email = string.IsNullOrWhiteSpace(EmailTextBox.Text) ? "(Not specified by user)" : EmailTextBox.Text.Trim();
                string subject = "QicBoqMapper Add-in Activation Code Request";
                
                string body = $"Hello Ahmed,\n\n" +
                              $"I would like to request an activation code for the QicBoqMapper Add-in.\n\n" +
                              $"--- Device Details ---\n" +
                              $"Machine ID: {Environment.MachineName} ({Environment.UserName})\n" +
                              $"Email Address: {email}\n";

                string mailtoUrl = $"mailto:engkhalaf7@gmail.com?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body)}";
                
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(mailtoUrl) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open email client: {ex.Message}\n\nYou can manually email engkhalaf7@gmail.com with your details.", "License Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
