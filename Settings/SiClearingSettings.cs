using System;
using System.IO;
using Newtonsoft.Json;

namespace SiClearing.Settings
{
    public class SiClearingSettings
    {
        public string SubjectFilter { get; set; } = "File SiClearing - SospesiTotaliProprio";
        public string SenderFilter { get; set; } = "";
        public string SaveFolder { get; set; } = @"C:\Temp\SiClearing\";
        public string SearchFolder { get; set; } = "";
        public string DefaultMarket { get; set; } = "ETFP";
        public string DefaultPortfolio { get; set; } = "ETF_EQUITY";
        public string DefaultColumns { get; set; } = "";

        private static readonly string SettingsPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "SiClearing", "settings.json");

        public static SiClearingSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    return JsonConvert.DeserializeObject<SiClearingSettings>(json)
                           ?? new SiClearingSettings();
                }
            }
            catch { }
            return new SiClearingSettings();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
                File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(this, Formatting.Indented));
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Errore salvataggio impostazioni:\n{ex.Message}",
                    "SiClearing", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }
    }
}
