using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using EnterpriseDocClassifier.Models;

namespace DocClassifier.AdminUI
{
    public partial class Form1 : Form
    {
        private ComboBox cbEnforcementMode, cbPlatform, cbPlacement, cbDefaultTag;
        private TextBox txtCustomMessage, txtCustomWarnMessage, txtName, txtMarkerText, txtColorHex;
        private NumericUpDown numFontSize;
        private DataGridView dgvTags;
        private PluginConfiguration _currentConfig;

        private readonly string _configDirectory = @"C:\ProgramData\YourCompany\DocClassifier";
        private readonly string _configPath = @"C:\ProgramData\YourCompany\DocClassifier\config.json";

        public Form1()
        {
            InitializeComponent();
            BuildModernUI();
            LoadExistingConfig();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        private void BuildModernUI()
        {
            this.Text = "Enterprise Data Loss Prevention - Dashboard";
            this.Size = new Size(950, 800);
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.BackColor = Color.FromArgb(243, 244, 246);
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(31, 41, 55) };
            pnlHeader.Controls.Add(new Label { Text = "DLP Policy Management Console", ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 16F), Location = new Point(20, 15), AutoSize = true });
            this.Controls.Add(pnlHeader);

            // --- SETTINGS CARD ---
            Panel cardSettings = CreateCard(20, 80, 890, 170);
            cardSettings.Controls.Add(new Label { Text = "Global Enforcement Settings", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(15, 10), AutoSize = true });

            cardSettings.Controls.Add(new Label { Text = "Enforcement Mode:", Location = new Point(15, 40), AutoSize = true, Font = new Font("Segoe UI Semibold", 10F) });
            cbEnforcementMode = new ComboBox { Location = new Point(180, 37), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cbEnforcementMode.Items.AddRange(new string[] { "None", "Warn", "Block" });
            cardSettings.Controls.Add(cbEnforcementMode);

            cardSettings.Controls.Add(new Label { Text = "Custom Block Message:", Location = new Point(15, 70), AutoSize = true });
            txtCustomMessage = new TextBox { Location = new Point(180, 67), Width = 680 };
            cardSettings.Controls.Add(txtCustomMessage);

            cardSettings.Controls.Add(new Label { Text = "Custom Warn Message:", Location = new Point(15, 100), AutoSize = true });
            txtCustomWarnMessage = new TextBox { Location = new Point(180, 97), Width = 680 };
            cardSettings.Controls.Add(txtCustomWarnMessage);

            cardSettings.Controls.Add(new Label { Text = "Default Tag (Auto-applied):", Location = new Point(15, 130), AutoSize = true });
            cbDefaultTag = new ComboBox { Location = new Point(180, 127), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            cardSettings.Controls.Add(cbDefaultTag);
            this.Controls.Add(cardSettings);

            // --- CREATE TAG CARD ---
            Panel cardCreate = CreateCard(20, 270, 890, 150);
            cardCreate.Controls.Add(new Label { Text = "Create New Sensitivity Tag", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(15, 10), AutoSize = true });

            cardCreate.Controls.Add(new Label { Text = "Platform:", Location = new Point(15, 45), AutoSize = true });
            cbPlatform = new ComboBox { Location = new Point(15, 65), Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
            cbPlatform.Items.AddRange(new string[] { "All", "Word", "Excel", "PowerPoint" });
            cbPlatform.SelectedIndex = 0;
            cardCreate.Controls.Add(cbPlatform);

            cardCreate.Controls.Add(new Label { Text = "Tag Name:", Location = new Point(140, 45), AutoSize = true });
            txtName = new TextBox { Location = new Point(140, 65), Width = 130 };
            cardCreate.Controls.Add(txtName);

            cardCreate.Controls.Add(new Label { Text = "Watermark Text:", Location = new Point(285, 45), AutoSize = true });
            txtMarkerText = new TextBox { Location = new Point(285, 65), Width = 200 };
            cardCreate.Controls.Add(txtMarkerText);

            cardCreate.Controls.Add(new Label { Text = "Placement:", Location = new Point(500, 45), AutoSize = true });
            cbPlacement = new ComboBox { Location = new Point(500, 65), Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
            cbPlacement.Items.AddRange(new string[] { "Top Left", "Top Center", "Top Right", "Bottom Left", "Bottom Center", "Bottom Right" });
            cbPlacement.SelectedIndex = 1;
            cardCreate.Controls.Add(cbPlacement);

            cardCreate.Controls.Add(new Label { Text = "Size:", Location = new Point(645, 45), AutoSize = true });
            numFontSize = new NumericUpDown { Location = new Point(645, 65), Width = 50, Minimum = 8, Maximum = 72, Value = 12 };
            cardCreate.Controls.Add(numFontSize);

            cardCreate.Controls.Add(new Label { Text = "Color:", Location = new Point(710, 45), AutoSize = true });
            txtColorHex = new TextBox { Location = new Point(710, 65), Width = 70, Text = "#FF0000", ReadOnly = true };
            cardCreate.Controls.Add(txtColorHex);

            Button btnPickColor = CreateFlatButton("...", 785, 64, 35, Color.FromArgb(107, 114, 128));
            btnPickColor.Click += BtnPickColor_Click;
            cardCreate.Controls.Add(btnPickColor);

            Button btnAdd = CreateFlatButton("Add Tag to Policy", 15, 100, 150, Color.FromArgb(37, 99, 235));
            btnAdd.Click += BtnAdd_Click;
            cardCreate.Controls.Add(btnAdd);
            this.Controls.Add(cardCreate);

            // --- DATAGRID CARD ---
            Panel cardGrid = CreateCard(20, 440, 890, 200);
            dgvTags = new DataGridView { Location = new Point(15, 15), Size = new Size(860, 170), AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, BackgroundColor = Color.White, BorderStyle = BorderStyle.None };
            dgvTags.Columns.Add("Platform", "Target Platform");
            dgvTags.Columns.Add("Name", "Tag Name");
            dgvTags.Columns.Add("Text", "Watermark Text");
            dgvTags.Columns.Add("Placement", "Placement");
            dgvTags.Columns.Add("Size", "Size");
            dgvTags.Columns.Add("Color", "Hex Color");
            cardGrid.Controls.Add(dgvTags);
            this.Controls.Add(cardGrid);

            Button btnRemove = CreateFlatButton("Remove Selected Tag", 20, 660, 180, Color.FromArgb(220, 38, 38));
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            Button btnSave = CreateFlatButton("Push Enterprise Policy", 660, 655, 250, Color.FromArgb(16, 185, 129));
            btnSave.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnSave.Height = 45;
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);
        }

        private Panel CreateCard(int x, int y, int w, int h)
        {
            return new Panel { Location = new Point(x, y), Size = new Size(w, h), BackColor = Color.White, Padding = new Padding(10) };
        }

        private Button CreateFlatButton(string text, int x, int y, int width, Color bgColor)
        {
            return new Button { Text = text, Location = new Point(x, y), Width = width, Height = 30, BackColor = bgColor, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        }

        private void LoadExistingConfig()
        {
            if (File.Exists(_configPath))
            {
                try { _currentConfig = JsonSerializer.Deserialize<PluginConfiguration>(File.ReadAllText(_configPath)) ?? new PluginConfiguration(); }
                catch { _currentConfig = new PluginConfiguration(); }
            }
            else { _currentConfig = new PluginConfiguration(); }

            if (_currentConfig.Classifications == null) _currentConfig.Classifications = new List<ClassificationLabel>();

            if (string.IsNullOrEmpty(_currentConfig.EnforcementMode)) _currentConfig.EnforcementMode = "None";

            cbEnforcementMode.SelectedItem = _currentConfig.EnforcementMode;
            txtCustomMessage.Text = _currentConfig.CustomBlockMessage ?? "";
            txtCustomWarnMessage.Text = _currentConfig.CustomWarnMessage ?? "";

            RefreshGrid();
        }

        private void RefreshGrid()
        {
            dgvTags.Rows.Clear();
            cbDefaultTag.Items.Clear();
            cbDefaultTag.Items.Add("None (Users must select manually)");

            foreach (var tag in _currentConfig.Classifications)
            {
                dgvTags.Rows.Add(tag.TargetPlatform, tag.Name, tag.Marker.Text, tag.Marker.Placement, tag.Marker.FontSize, tag.Marker.FontColor);
                cbDefaultTag.Items.Add(tag.Name);
            }

            if (!string.IsNullOrEmpty(_currentConfig.DefaultClassificationName) && cbDefaultTag.Items.Contains(_currentConfig.DefaultClassificationName))
                cbDefaultTag.SelectedItem = _currentConfig.DefaultClassificationName;
            else
                cbDefaultTag.SelectedIndex = 0;
        }

        private void BtnPickColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog() == DialogResult.OK) txtColorHex.Text = ColorTranslator.ToHtml(Color.FromArgb(cd.Color.ToArgb()));
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtMarkerText.Text)) return;

            _currentConfig.Classifications.Add(new ClassificationLabel
            {
                Name = txtName.Text.Trim(),
                TargetPlatform = cbPlatform.Text,
                Marker = new DocumentMarker { Text = txtMarkerText.Text.Trim(), Placement = cbPlacement.Text, FontSize = (int)numFontSize.Value, FontColor = txtColorHex.Text }
            });
            RefreshGrid(); txtName.Clear(); txtMarkerText.Clear();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (dgvTags.SelectedRows.Count > 0)
            {
                _currentConfig.Classifications.RemoveAt(dgvTags.SelectedRows[0].Index);
                RefreshGrid();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            _currentConfig.EnforcementMode = cbEnforcementMode.SelectedItem.ToString();
            _currentConfig.CustomBlockMessage = txtCustomMessage.Text.Trim();
            _currentConfig.CustomWarnMessage = txtCustomWarnMessage.Text.Trim();
            _currentConfig.DefaultClassificationName = cbDefaultTag.SelectedIndex <= 0 ? "" : cbDefaultTag.Text;

            try
            {
                if (!Directory.Exists(_configDirectory)) Directory.CreateDirectory(_configDirectory);
                File.WriteAllText(_configPath, JsonSerializer.Serialize(_currentConfig, new JsonSerializerOptions { WriteIndented = true }));
                MessageBox.Show("Policy saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Error: {ex.Message}"); }
        }
    }
}