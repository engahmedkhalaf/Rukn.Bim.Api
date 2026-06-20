using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace QicBoqMapper.Helpers
{
    public static class IconHelper
    {
        public static BitmapImage? LoadIcon(string iconName)
        {
            try
            {
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string assemblyDirectory = Path.GetDirectoryName(assemblyLocation) ?? string.Empty;
                string iconPath = Path.Combine(assemblyDirectory, "Resources", iconName);

                if (!File.Exists(iconPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[Warning] Icon not found: {iconPath}");
                    return null;
                }

                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(iconPath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                return bitmap;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Warning] Error loading icon {iconName}: {ex.Message}");
                return null;
            }
        }
    }
}
