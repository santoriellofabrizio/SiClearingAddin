using ExcelDna.Integration;
using SiClearing.Settings;

namespace SiClearing
{
    public class AddIn : IExcelAddIn
    {
        internal static SiClearingSettings Settings { get; private set; } = SiClearingSettings.Load();
        internal static CsvCache Cache { get; } = new CsvCache();

        public void AutoOpen()
        {
            Settings = SiClearingSettings.Load();
        }

        public void AutoClose()
        {
        }

        public static void InvalidateCache()
        {
            Cache.Invalidate();
        }

        public static void ReloadSettings()
        {
            Settings = SiClearingSettings.Load();
        }
    }
}
