using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SiClearing
{
    [ComVisible(true)]
    public class SiClearingRibbon : ExcelRibbon
    {
        private IRibbonUI? _ribbon;

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void OnRefresh(IRibbonControl control)
        {
            try
            {
                var downloader = new OutlookDownloader();
                string? path = downloader.DownloadLatestToday(AddIn.Settings);

                if (path != null)
                {
                    AddIn.Cache.Invalidate();
                    AddIn.Cache.LoadLatest(AddIn.Settings.SaveFolder);
                    MessageBox.Show($"CSV scaricato e caricato:\n{path}",
                        "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ExcelDna.Integration.XlCall.Excel(ExcelDna.Integration.XlCall.xlcCalculateNow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante Refresh:\n{ex.Message}",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSettings(IRibbonControl control)
        {
            using var form = new SettingsForm(AddIn.Settings);
            form.ShowDialog();
            AddIn.ReloadSettings();
        }

        public void OnDebugFolders(IRibbonControl control)
        {
            try
            {
                Outlook.Application? outlook;
                try
                {
                    outlook = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    MessageBox.Show("Outlook non è aperto.", "SiClearing",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var ns = outlook.GetNamespace("MAPI");
                var xl = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                var sheet = (Microsoft.Office.Interop.Excel.Worksheet)xl.Sheets.Add();
                sheet.Name = "OutlookFolders";

                var downloader = new OutlookDownloader();
                downloader.DebugFolders(ns, sheet);

                MessageBox.Show("Struttura cartelle Outlook scritta nel foglio 'OutlookFolders'.",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore debug cartelle:\n{ex.Message}",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
