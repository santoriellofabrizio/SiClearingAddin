using System;
using System.Drawing;
using System.Windows.Forms;

namespace SiClearing
{
    public class LogForm : Form
    {
        private readonly RichTextBox _box;
        private readonly Button _btnClear;
        private readonly Button _btnCopy;

        private static LogForm? _instance;

        public static void ShowOrActivate()
        {
            if (_instance == null || _instance.IsDisposed)
            {
                _instance = new LogForm();
                _instance.Show();
            }
            else
            {
                _instance.Activate();
                if (_instance.WindowState == FormWindowState.Minimized)
                    _instance.WindowState = FormWindowState.Normal;
            }
        }

        private LogForm()
        {
            Text = "SiClearing — Log";
            Width = 700;
            Height = 420;
            MinimumSize = new Size(400, 200);
            StartPosition = FormStartPosition.Manual;
            Location = new Point(100, 100);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;

            _box = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9f),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.LightGreen,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = false,
            };

            _btnClear = new Button { Text = "Pulisci", Width = 80, Dock = DockStyle.Right };
            _btnCopy  = new Button { Text = "Copia tutto", Width = 90, Dock = DockStyle.Right };
            _btnClear.Click += (s, e) => { Logger.Clear(); _box.Clear(); };
            _btnCopy.Click  += (s, e) => { if (_box.Text.Length > 0) Clipboard.SetText(_box.Text); };

            var toolbar = new Panel { Dock = DockStyle.Bottom, Height = 32 };
            toolbar.Controls.Add(_btnClear);
            toolbar.Controls.Add(_btnCopy);

            Controls.Add(_box);
            Controls.Add(toolbar);

            // populate existing entries
            foreach (var e in Logger.Entries)
                AppendEntry(e);

            Logger.EntryAdded += OnEntryAdded;
            FormClosed += (s, e) => Logger.EntryAdded -= OnEntryAdded;
        }

        private void OnEntryAdded(string entry)
        {
            if (InvokeRequired)
                BeginInvoke(new Action<string>(AppendEntry), entry);
            else
                AppendEntry(entry);
        }

        private void AppendEntry(string entry)
        {
            _box.AppendText(entry + Environment.NewLine);
            _box.ScrollToCaret();
        }
    }
}
