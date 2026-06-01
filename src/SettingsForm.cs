using System.Windows.Forms;
using SiClearing.Settings;

namespace SiClearing
{
    public class SettingsForm : Form
    {
        private readonly TextBox _txtSubject = new TextBox();
        private readonly TextBox _txtSender = new TextBox();
        private readonly TextBox _txtSaveFolder = new TextBox();
        private readonly TextBox _txtSearchFolder = new TextBox();
        private readonly TextBox _txtMarket = new TextBox();
        private readonly TextBox _txtPortfolio = new TextBox();
        private readonly TextBox _txtDefaultColumns = new TextBox();
        private readonly TextBox _txtMercatoFilter = new TextBox();
        private readonly TextBox _txtTipoContoFilter = new TextBox();
        private readonly SiClearingSettings _settings;

        public SettingsForm(SiClearingSettings settings)
        {
            _settings = settings;
            Text = "Impostazioni SiClearing";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Width = 520;
            AutoSize = true;
            Padding = new Padding(10);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
                Padding = new Padding(10)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

            AddRow(layout, "Filtro oggetto mail:", _txtSubject, settings.SubjectFilter, null);
            AddRow(layout, "Filtro mittente:", _txtSender, settings.SenderFilter, null);
            AddRow(layout, "Cartella salvataggio:", _txtSaveFolder, settings.SaveFolder, BrowseFolder);
            AddRow(layout, "Sottocartella Outlook:", _txtSearchFolder, settings.SearchFolder, null);
            AddRow(layout, "Mercato default:", _txtMarket, settings.DefaultMarket, null);
            AddRow(layout, "Portfolio default:", _txtPortfolio, settings.DefaultPortfolio, null);
            AddRow(layout, "Colonne default (sep. ;):", _txtDefaultColumns, settings.DefaultColumns, null);
            AddRow(layout, "Filtro Mercato (sep. ;):", _txtMercatoFilter, settings.MercatoFilter, null);
            AddRow(layout, "Filtro Tipo Conto (sep. ;):", _txtTipoContoFilter, settings.TipoContoFilter, null);

            var btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
            var btnCancel = new Button { Text = "Annulla", DialogResult = DialogResult.Cancel, Width = 80 };
            btnOk.Click += (s, e) => Save();

            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true
            };
            btnPanel.Controls.AddRange(new Control[] { btnCancel, btnOk });

            Controls.Add(layout);
            Controls.Add(btnPanel);
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void AddRow(TableLayoutPanel layout, string label, TextBox box, string value, System.Action? browseAction)
        {
            box.Text = value;
            box.Dock = DockStyle.Fill;
            box.Width = 250;

            layout.Controls.Add(new Label { Text = label, TextAlign = System.Drawing.ContentAlignment.MiddleRight, AutoSize = true });
            layout.Controls.Add(box);

            if (browseAction != null)
            {
                var btn = new Button { Text = "Sfoglia...", Width = 80 };
                btn.Click += (s, e) => browseAction();
                layout.Controls.Add(btn);
            }
            else
            {
                layout.Controls.Add(new Label()); // placeholder
            }
        }

        private void BrowseFolder()
        {
            using var dlg = new FolderBrowserDialog { SelectedPath = _txtSaveFolder.Text };
            if (dlg.ShowDialog() == DialogResult.OK)
                _txtSaveFolder.Text = dlg.SelectedPath + System.IO.Path.DirectorySeparatorChar;
        }

        private void Save()
        {
            bool folderChanged = _settings.SaveFolder != _txtSaveFolder.Text;

            _settings.SubjectFilter = _txtSubject.Text;
            _settings.SenderFilter = _txtSender.Text;
            _settings.SaveFolder = _txtSaveFolder.Text;
            _settings.SearchFolder = _txtSearchFolder.Text;
            _settings.DefaultMarket = _txtMarket.Text;
            _settings.DefaultPortfolio = _txtPortfolio.Text;
            _settings.DefaultColumns = _txtDefaultColumns.Text;
            _settings.MercatoFilter = _txtMercatoFilter.Text;
            _settings.TipoContoFilter = _txtTipoContoFilter.Text;
            _settings.Save();

            if (folderChanged)
                AddIn.InvalidateCache();
        }
    }
}
