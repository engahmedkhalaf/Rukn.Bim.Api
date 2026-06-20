using System;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace QicBoqMapper
{
    public static class LicenseManager
    {
        private const string RegistryPath = @"Software\QicTools\QicBoqMapper";
        private const string EmailValueName = "Email";
        private const string CodeValueName = "ActivationCode";
        private const string LastVerifiedValueName = "LastVerified";
        private const string ExpiresAtValueName = "ExpiresAt";

        // Supabase configuration placeholders
        private const string SupabaseUrl = "https://auvtapbsdewwmzejchgq.supabase.co";
        private const string SupabaseAnonKey = "sb_publishable_LUcQW4gZYVFYGsNY-p-n0A_LSr2N4RH";

        public static bool IsActivated()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key == null) return false;
                    
                    object lastVerifiedVal = key.GetValue(LastVerifiedValueName);
                    if (lastVerifiedVal == null || string.IsNullOrWhiteSpace(lastVerifiedVal.ToString()))
                        return false;

                    // Check local expiration date
                    object expiresAtVal = key.GetValue(ExpiresAtValueName);
                    if (expiresAtVal != null && !string.IsNullOrWhiteSpace(expiresAtVal.ToString()))
                    {
                        if (DateTimeOffset.TryParse(expiresAtVal.ToString(), out DateTimeOffset expiresAt))
                        {
                            if (DateTimeOffset.UtcNow > expiresAt)
                            {
                                return false; // Expired locally
                            }
                        }
                    }

                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public static string GetSavedEmail()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key == null) return string.Empty;
                    return key.GetValue(EmailValueName)?.ToString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string GetSavedCode()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key == null) return string.Empty;
                    return key.GetValue(CodeValueName)?.ToString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static void SaveLicense(string email, string code, bool isActivated, string expiresAtStr = null)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    if (key != null)
                    {
                        key.SetValue(EmailValueName, email ?? string.Empty);
                        key.SetValue(CodeValueName, code ?? string.Empty);
                        
                        if (isActivated)
                        {
                            string utcNowStr = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.FFFFFFFZ");
                            key.SetValue(LastVerifiedValueName, utcNowStr, RegistryValueKind.String);
                            
                            if (!string.IsNullOrWhiteSpace(expiresAtStr))
                            {
                                key.SetValue(ExpiresAtValueName, expiresAtStr, RegistryValueKind.String);
                            }
                            else
                            {
                                try { key.DeleteValue(ExpiresAtValueName, false); } catch { }
                            }
                        }
                        else
                        {
                            try { key.DeleteValue(LastVerifiedValueName, false); } catch { }
                            try { key.DeleteValue(ExpiresAtValueName, false); } catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save license in Registry: {ex.Message}");
            }
        }

        public static bool ValidateInput(string email, string code)
        {
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(code))
                return false;

            bool validEmail = Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
            bool validCode = code.Trim().Length >= 4;

            return validEmail && validCode;
        }

        public static async Task<Tuple<bool, string>> ValidateLicenseWithSupabaseAsync(string email, string code)
        {
            if (!ValidateInput(email, code))
                return Tuple.Create(false, (string)null);

            try
            {
                string cleanEmail = email.Trim().ToLower();
                string cleanCode = code.Trim();

                string encodedEmail = Uri.EscapeDataString(cleanEmail);
                string encodedCode = Uri.EscapeDataString(cleanCode);
                
                // Use 'ilike' for case-insensitive email match, and exact 'eq' for the activation code
                string requestUrl = $"{SupabaseUrl.TrimEnd('/')}/rest/v1/licenses?email=ilike.{encodedEmail}&activation_code=eq.{encodedCode}&select=*";

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("apikey", SupabaseAnonKey);
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {SupabaseAnonKey}");
                    
                    var response = await client.GetAsync(requestUrl).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        throw new Exception($"Supabase server returned status code: {response.StatusCode}. Details: {errorContent}");
                    }

                    string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    string trimmedResponse = jsonResponse.Trim();
                    
                    if (!trimmedResponse.StartsWith("[") || !trimmedResponse.EndsWith("]"))
                    {
                        return Tuple.Create(false, (string)null);
                    }

                    if (trimmedResponse == "[]")
                    {
                        return Tuple.Create(false, (string)null);
                    }

                    if (trimmedResponse.Contains("\"status\"") && !trimmedResponse.Contains("\"status\":\"active\"") && !trimmedResponse.Contains("\"status\": \"active\""))
                    {
                        return Tuple.Create(false, (string)null);
                    }

                    // Check for license expiration date if present
                    string expiresPattern = "\"expires_at\"\\s*:\\s*\"([^\"]+)\"";
                    var match = Regex.Match(trimmedResponse, expiresPattern);
                    string expiresAtStr = null;
                    if (match.Success)
                    {
                        expiresAtStr = match.Groups[1].Value;
                        if (DateTimeOffset.TryParse(expiresAtStr, out DateTimeOffset expiresAt))
                        {
                            if (DateTimeOffset.UtcNow > expiresAt)
                            {
                                return Tuple.Create(false, (string)null); // License has expired
                            }
                        }
                    }

                    return Tuple.Create(true, expiresAtStr);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Supabase connection failed: {ex.Message}");
            }
        }
    }
}
