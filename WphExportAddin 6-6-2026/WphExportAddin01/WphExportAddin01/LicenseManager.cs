using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace WphExportAddin
{
    /// <summary>
    /// Manages WPH Export Add-in licensing.
    /// Saves and checks license parameters in the Windows Registry (HKCU\Software\WPH\Export).
    /// Performs silent online validation with a 7-day offline fallback grace period.
    /// </summary>
    public static class LicenseManager
    {
        // =====================================================================
        // CONFIGURATION: Supabase credentials.
        // =====================================================================
        private const string SupabaseUrl = "https://dfkcnyzuiquvozvncwph.supabase.co";
        private const string SupabaseAnonKey = "sb_publishable_zhW-Ox8_ssRAZKkGkBbsog_1juWTr1X";
        // =====================================================================

        private const string RegistryKeyPath = @"Software\WPH\Export";

        /// <summary>
        /// Checks if a valid license exists.
        /// Performs a quick online check; if offline, falls back to a 7-day grace period.
        /// </summary>
        public static bool IsLicensed()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key == null)
                        return false;

                    string email = key.GetValue("Email")?.ToString();
                    string code = key.GetValue("ActivationCode")?.ToString();
                    string lastVerifiedStr = key.GetValue("LastVerified")?.ToString();

                    if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(code))
                        return false;

                    // Master / Offline bypass keys require no online validation
                    if (code == "WPH-MASTER-2026-TRUSTED" || code.StartsWith("WPH-LOCAL-"))
                        return true;

                    if (!DateTime.TryParse(lastVerifiedStr, out DateTime lastVerified))
                        return false;

                    // Attempt a quick online validation (3 second timeout)
                    if (Activate(email, code, out string error, 3))
                    {
                        return true; // Still active and valid
                    }
                    else
                    {
                        // If it failed because the network/connection is down, allow 7-day grace period
                        if (error != null && error.StartsWith("Network connection failed"))
                        {
                            double days = (DateTime.UtcNow - lastVerified.ToUniversalTime()).TotalDays;
                            if (days >= 0 && days <= 7.0)
                            {
                                return true; // Allowed inside 7-day offline grace period
                            }
                        }

                        // Otherwise, the database returned that the license is invalid/deleted.
                        // Wipe registry and block access.
                        DeleteLicenseLocally();
                        return false;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Attempts to activate the product. Saves credentials directly to HKCU\Software\WPH\Export upon success.
        /// </summary>
        public static bool Activate(string email, string code, out string errorMessage, int timeoutSeconds = 10)
        {
            errorMessage = string.Empty;

            if (string.IsNullOrWhiteSpace(email) || !email.Contains("@"))
            {
                errorMessage = "Please enter a valid email address.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(code))
            {
                errorMessage = "Please enter an activation code.";
                return false;
            }

            email = email.Trim();
            code = code.Trim();

            // 1. Master Bypass Codes (works offline)
            if (code == "WPH-MASTER-2026-TRUSTED" || code.StartsWith("WPH-LOCAL-"))
            {
                SaveLicenseLocally(email, code);
                return true;
            }

            // Ensure modern TLS protocol is used
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= 
                    System.Net.SecurityProtocolType.Tls12 | (System.Net.SecurityProtocolType)3072;
            }
            catch { }

            // 2. Query Supabase
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(timeoutSeconds);

                    // Add required Supabase headers
                    client.DefaultRequestHeaders.Add("apikey", SupabaseAnonKey);
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + SupabaseAnonKey);

                    // Build query URL
                    string queryUrl = string.Format(
                        "{0}/rest/v1/licenses?email=eq.{1}&activation_code=eq.{2}",
                        SupabaseUrl.TrimEnd('/'),
                        Uri.EscapeDataString(email),
                        Uri.EscapeDataString(code)
                    );

                    HttpResponseMessage response = Task.Run(() => client.GetAsync(queryUrl)).GetAwaiter().GetResult();

                    if (!response.IsSuccessStatusCode)
                    {
                        errorMessage = string.Format("Server returned error status: {0} ({1})", 
                            (int)response.StatusCode, response.ReasonPhrase);
                        return false;
                    }

                    string responseBody = Task.Run(() => response.Content.ReadAsStringAsync()).GetAwaiter().GetResult();

                    bool hasLicense = !string.IsNullOrEmpty(responseBody) && responseBody.Length > 2;

                    if (!hasLicense)
                    {
                        errorMessage = "Invalid email or activation code. Please check your credentials.";
                        return false;
                    }

                    SaveLicenseLocally(email, code);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Exception root = ex;
                while (root.InnerException != null) root = root.InnerException;
                errorMessage = "Network connection failed. Use offline activation code or try again: " + root.Message;
                return false;
            }
        }

        /// <summary>
        /// Saves the active license variables in the Registry under HKCU\Software\WPH\Export.
        /// </summary>
        private static void SaveLicenseLocally(string email, string code)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("Email", email);
                        key.SetValue("ActivationCode", code);
                        key.SetValue("LastVerified", DateTime.UtcNow.ToString("o"));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Could not save the license to the Windows Registry: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Deletes the local registry key values to deactivate the client.
        /// </summary>
        public static void DeleteLicenseLocally()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(RegistryKeyPath, false);
            }
            catch { }
        }

        /// <summary>
        /// Gets a basic device identifier string.
        /// </summary>
        public static string GetMachineIdentifier()
        {
            return string.Format("{0} ({1})", Environment.MachineName, Environment.UserName);
        }
    }
}
