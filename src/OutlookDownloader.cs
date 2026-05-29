using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SiClearing.Settings;

namespace SiClearing
{
    public class OutlookDownloader
    {
        public string? DownloadLatestToday(SiClearingSettings settings)
        {
            Application? outlook = null;
            try
            {
                try
                {
                    outlook = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
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

                foreach (Attachment att in mail.Attachments)
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
            catch (System.Exception ex)
            {
                MessageBox.Show($"Errore durante il download:\n{ex.Message}",
                    "SiClearing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private MAPIFolder ResolveFolder(NameSpace ns, string folderPath)
        {
            var inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            if (string.IsNullOrWhiteSpace(folderPath))
                return inbox;

            var parts = folderPath.Split('\\');
            MAPIFolder current = inbox;
            foreach (var part in parts)
            {
                bool found = false;
                foreach (MAPIFolder sub in current.Folders)
                {
                    if (sub.Name.Equals(part, StringComparison.OrdinalIgnoreCase))
                    {
                        current = sub;
                        found = true;
                        break;
                    }
                }
                if (!found)
                    throw new System.Exception($"Cartella Outlook non trovata: '{part}' in '{current.Name}'");
            }
            return current;
        }

        private MailItem? SearchFolderRestrict(MAPIFolder folder, SiClearingSettings settings)
        {
            string filter = $"[ReceivedTime] >= '{DateTime.Today:MM/dd/yyyy} 00:00 AM' " +
                            $"AND [ReceivedTime] <= '{DateTime.Today:MM/dd/yyyy} 11:59 PM'";

            var items = folder.Items.Restrict(filter);
            items.Sort("[ReceivedTime]", true);

            MailItem? best = null;

            for (int i = 1; i <= items.Count; i++)
            {
                var obj = items.Item(i);
                if (!(obj is MailItem mail)) continue;

                if (!mail.Subject.Contains(settings.SubjectFilter)) continue;

                if (!string.IsNullOrWhiteSpace(settings.SenderFilter) &&
                    !mail.SenderEmailAddress.Contains(settings.SenderFilter))
                    continue;

                best = mail;
                break; // items are sorted descending by ReceivedTime
            }

            return best;
        }

        public void DebugFolders(NameSpace ns, Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            int row = 1;
            WriteFolder(ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox), sheet, ref row, 0);
        }

        private void WriteFolder(MAPIFolder folder, Microsoft.Office.Interop.Excel.Worksheet sheet,
            ref int row, int depth)
        {
            sheet.Cells[row, 1] = new string(' ', depth * 2) + folder.Name;
            row++;
            foreach (MAPIFolder sub in folder.Folders)
                WriteFolder(sub, sheet, ref row, depth + 1);
        }
    }
}
