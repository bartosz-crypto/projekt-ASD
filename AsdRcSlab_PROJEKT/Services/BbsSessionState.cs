using System;
using System.IO;

namespace AsdRcSlab
{
    /// <summary>
    /// Stan sesji BBS z persistent storage między uruchomieniami AutoCAD.
    /// Plik konfiguracyjny: %APPDATA%\AsdRcSlab\settings.json
    /// </summary>
    public static class BbsSessionState
    {
        private static string _lastTemplatePath;
        private static bool _loaded;

        public static string LastTemplatePath
        {
            get
            {
                EnsureLoaded();
                return _lastTemplatePath;
            }
            set
            {
                EnsureLoaded();
                if (_lastTemplatePath == value) return;
                _lastTemplatePath = value;
                SaveToDisk();
            }
        }

        private static string GetSettingsPath()
        {
            string appData = Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData);
            string dir = Path.Combine(appData, "AsdRcSlab");
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return Path.Combine(dir, "settings.json");
        }

        private static void EnsureLoaded()
        {
            if (_loaded) return;
            _loaded = true;
            try
            {
                string path = GetSettingsPath();
                if (!File.Exists(path)) return;
                string json = File.ReadAllText(path);
                // Prosty parsing — szukamy "LastTemplatePath":"..."
                // Bez bibliotek zewnętrznych. Wystarcza dla 1 klucza.
                var match = System.Text.RegularExpressions.Regex.Match(
                    json,
                    @"""LastTemplatePath""\s*:\s*""([^""]*)""");
                if (match.Success)
                {
                    string val = match.Groups[1].Value
                        .Replace(@"\\", @"\")
                        .Replace(@"\""", @"""");
                    _lastTemplatePath = val;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[BbsSessionState] Load failed: " + ex.Message);
            }
        }

        private static void SaveToDisk()
        {
            try
            {
                string path = GetSettingsPath();
                string escaped = (_lastTemplatePath ?? "")
                    .Replace(@"\", @"\\")
                    .Replace(@"""", @"\""");
                string json = "{\n  \"LastTemplatePath\": \""
                    + escaped + "\"\n}\n";
                File.WriteAllText(path, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[BbsSessionState] Save failed: " + ex.Message);
            }
        }
    }
}
