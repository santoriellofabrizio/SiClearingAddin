using System;
using System.IO;
using System.Windows.Forms;
using SiClearing.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SiClearing
{
    public class OutlookDownloader
    {
        public string? DownloadLatestToday(SiClearingSettings settings)
        {
            Outlook.Application? outlook = null;
            try
            {
                try
                {
                    outlook = (Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    MessageBox.Show("Outlook non è aperto. Aprire Outlook e riprovare.",
                        "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                var ns = outlook.GetNamespace("MAPI");
                var folder = ResolveFolder(ns, settings.SearchFolder);
                var mail = SearchFolderRestrict(folder, settings);

                if (mail == null)
                {
                    MessageBox.Show("Nessuna mail trovata oggi con il filtro specificato.",
                        "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                Directory.CreateDirectory(settings.SaveFolder);

                foreach (Outlook.Attachment att in mail.Attachments)
                {
                    if (att.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        string dest = Path.Combine(settings.SaveFolder,
                            $"{DateTime.Today:yyyyMMdd}_{att.FileName}");
                        att.SaveAsFile(dest);
                        return dest;
                    }
                }

                MessageBox.Show("Mail trovata ma nessun allegato CSV.",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante il download:\n{ex.Message}",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private Outlook.MAPIFolder ResolveFolder(Outlook.NameSpace ns, string folderPath)
        {
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            if (string.IsNullOrWhiteSpace(folderPath))
                return inbox;

            var parts = folderPath.Split('\\');
            Outlook.MAPIFolder current = inbox;
            foreach (var part in parts)
            {
                bool found = false;
                foreach (Outlook.MAPIFolder sub in current.Folders)
                {
                    if (sub.Name.Equals(part, StringComparison.OrdinalIgnoreCase))
                    {
                        current = sub;
                        found = true;
                        break;
                    }
                }
                if (!found)
                    throw new Exception($"Cartella Outlook non trovata: '{part}' in '{current.Name}'");
            }
            return current;
        }

        private Outlook.MailItem? SearchFolderRestrict(Outlook.MAPIFolder folder, SiClearingSettings settings)
        {
            string filter = $"[ReceivedTime] >= '{DateTime.Today:MM/dd/yyyy} 00:00 AM' " +
                            $"AND [ReceivedTime] <= '{DateTime.Today:MM/dd/yyyy} 11:59 PM'";

            var items = (Outlook.Items)folder.Items.Restrict(filter);
            items.Sort("[ReceivedTime]", true);

            for (int i = 1; i <= items.Count; i++)
            {
                var obj = items.Item((object)i);
                if (!(obj is Outlook.MailItem mail)) continue;

                if (!mail.Subject.Contains(settings.SubjectFilter)) continue;

                if (!string.IsNullOrWhiteSpace(settings.SenderFilter) &&
                    !mail.SenderEmailAddress.Contains(settings.SenderFilter))
                    continue;

                return mail;
            }

            return null;
        }

        public void DebugFolders(Outlook.NameSpace ns, Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            int row = 1;
            WriteFolder(ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), sheet, ref row, 0);
        }

        private void WriteFolder(Outlook.MAPIFolder folder, Microsoft.Office.Interop.Excel.Worksheet sheet,
            ref int row, int depth)
        {
            sheet.Cells[row, 1] = new string(' ', depth * 2) + folder.Name;
            row++;
            foreach (Outlook.MAPIFolder sub in folder.Folders)
                WriteFolder(sub, sheet, ref row, depth + 1);
        }
    }
}
