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
        // Define colors from your XAML mock
        private readonly Color AppBg = ColorTranslator.FromHtml("#F5F7FA");
        private readonly Color CardBg = ColorTranslator.FromHtml("#FFFFFF");
        private readonly Color TextPrimary = ColorTranslator.FromHtml("#0F172A");
        private readonly Color TextSecondary = ColorTranslator.FromHtml("#64748B");
        private readonly Color SidebarBg = ColorTranslator.FromHtml("#0B1220");
        private readonly Color PrimaryBlue = ColorTranslator.FromHtml("#2563EB");
        private readonly Color SuccessGreen = ColorTranslator.FromHtml("#16A34A");
        private readonly Color DangerRed = ColorTranslator.FromHtml("#EF4444");
        private readonly Color BorderSoft = ColorTranslator.FromHtml("#E5EAF0");

        private ComboBox cbEnforcementMode, cbPlatform, cbPlacement, cbDefaultTag;
        private TextBox txtCustomMessage, txtCustomWarnMessage, txtName, txtMarkerText, txtColorHex;
        private NumericUpDown numFontSize;
        private DataGridView dgvTags;
        private Panel colorPreview;
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
            this.Text = "Enterprise DLP Management";
            this.Size = new Size(1400, 900);
            this.BackColor = AppBg;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 1. Scrollable Main Content (Added first so Fill works correctly)
            Panel mainContent = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = AppBg };
            this.Controls.Add(mainContent);

            // 2. Top App Bar
            Panel topbar = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = CardBg };
            topbar.Paint += (s, e) => { ControlPaint.DrawBorder(e.Graphics, topbar.ClientRectangle, BorderSoft, ButtonBorderStyle.Solid); };
            topbar.Controls.Add(new Label { Text = "Enterprise DLP Management", Font = new Font("Segoe UI Semibold", 16), ForeColor = TextPrimary, Location = new Point(20, 16), AutoSize = true });
            this.Controls.Add(topbar);

            // 3. Left Nav Sidebar
            Panel sidebar = new Panel { Dock = DockStyle.Left, Width = 70, BackColor = SidebarBg };
            sidebar.Controls.Add(CreateIconLabel("🏠", 20));
            sidebar.Controls.Add(CreateIconLabel("📄", 70));
            sidebar.Controls.Add(CreateIconLabel("⚙", 120));
            sidebar.Controls.Add(CreateIconLabel("ℹ", 170));
            this.Controls.Add(sidebar);

            // --- MAIN CONTENT ELEMENTS ---

            // Hero Banner
            Panel hero = new Panel { Location = new Point(24, 24), Size = new Size(1300, 100), BackColor = SidebarBg };
            hero.Controls.Add(new Label { Text = "Enterprise DLP Management", Font = new Font("Segoe UI Semibold", 22), ForeColor = Color.White, Location = new Point(20, 20), AutoSize = true });
            hero.Controls.Add(new Label { Text = "Configure organization-wide sensitivity labels and enforcement policies.", Font = new Font("Segoe UI", 11), ForeColor = ColorTranslator.FromHtml("#C7D2FE"), Location = new Point(24, 62), AutoSize = true });
            mainContent.Controls.Add(hero);

            // Global Enforcement Settings Card
            Panel cardSettings = CreateCard(24, 140, 600, 380);
            cardSettings.Controls.Add(new Label { Text = "⚙ Global Enforcement Settings", Font = new Font("Segoe UI Semibold", 16), ForeColor = TextPrimary, Location = new Point(20, 20), AutoSize = true });
            cardSettings.Controls.Add(new Label { Text = "Manage default data loss prevention settings for the entire organization.", Font = new Font("Segoe UI", 10), ForeColor = TextSecondary, Location = new Point(22, 55), AutoSize = true });

            int startY = 110; int spacing = 55;
            cardSettings.Controls.Add(CreateFieldLabel("Enforcement Mode:", 20, startY));
            cbEnforcementMode = CreateComboBox(200, startY - 4, 350, new[] { "None", "Warn", "Block" });
            cardSettings.Controls.Add(cbEnforcementMode);

            cardSettings.Controls.Add(CreateFieldLabel("Default Tag:", 20, startY + spacing));
            cbDefaultTag = CreateComboBox(200, startY + spacing - 4, 350, new string[0]);
            cardSettings.Controls.Add(cbDefaultTag);

            cardSettings.Controls.Add(CreateFieldLabel("Custom Block Msg:", 20, startY + spacing * 2));
            txtCustomMessage = CreateTextBox(200, startY + spacing * 2 - 4, 350);
            cardSettings.Controls.Add(txtCustomMessage);

            cardSettings.Controls.Add(CreateFieldLabel("Custom Warn Msg:", 20, startY + spacing * 3));
            txtCustomWarnMessage = CreateTextBox(200, startY + spacing * 3 - 4, 350);
            cardSettings.Controls.Add(txtCustomWarnMessage);
            mainContent.Controls.Add(cardSettings);

            // Create Tag Card
            Panel cardCreate = CreateCard(640, 140, 684, 380);
            cardCreate.Controls.Add(new Label { Text = "🏷 Create New Sensitivity Tag", Font = new Font("Segoe UI Semibold", 16), ForeColor = TextPrimary, Location = new Point(20, 20), AutoSize = true });
            cardCreate.Controls.Add(new Label { Text = "Create a label and define how it should appear as a watermark.", Font = new Font("Segoe UI", 10), ForeColor = TextSecondary, Location = new Point(22, 55), AutoSize = true });

            cardCreate.Controls.Add(CreateFieldLabel("Platform:", 20, startY));
            cbPlatform = CreateComboBox(150, startY - 4, 150, new[] { "All", "Word", "Excel", "PowerPoint" });
            cardCreate.Controls.Add(cbPlatform);

            cardCreate.Controls.Add(CreateFieldLabel("Tag Name:", 20, startY + spacing));
            txtName = CreateTextBox(150, startY + spacing - 4, 200);
            cardCreate.Controls.Add(txtName);

            cardCreate.Controls.Add(CreateFieldLabel("Watermark Text:", 20, startY + spacing * 2));
            txtMarkerText = CreateTextBox(150, startY + spacing * 2 - 4, 280);
            cardCreate.Controls.Add(txtMarkerText);

            cardCreate.Controls.Add(CreateFieldLabel("Placement:", 20, startY + spacing * 3));
            cbPlacement = CreateComboBox(150, startY + spacing * 3 - 4, 180, new[] { "Top Left", "Top Center", "Top Right", "Bottom Left", "Bottom Center", "Bottom Right" });
            cardCreate.Controls.Add(cbPlacement);

            cardCreate.Controls.Add(CreateFieldLabel("Size:", 20, startY + spacing * 4));
            numFontSize = new NumericUpDown { Location = new Point(150, startY + spacing * 4 - 4), Width = 60, Minimum = 8, Maximum = 72, Value = 12, Font = new Font("Segoe UI", 11) };
            cardCreate.Controls.Add(numFontSize);

            cardCreate.Controls.Add(CreateFieldLabel("Hex Color:", 240, startY + spacing * 4));
            txtColorHex = CreateTextBox(325, startY + spacing * 4 - 4, 90);
            txtColorHex.Text = "#FF0000";
            cardCreate.Controls.Add(txtColorHex);

            colorPreview = new Panel { Location = new Point(425, startY + spacing * 4 - 4), Size = new Size(28, 28), BackColor = Color.Red, Cursor = Cursors.Hand };
            colorPreview.Paint += (s, e) => { ControlPaint.DrawBorder(e.Graphics, colorPreview.ClientRectangle, BorderSoft, ButtonBorderStyle.Solid); };
            colorPreview.Click += BtnPickColor_Click;
            cardCreate.Controls.Add(colorPreview);

            Button btnAdd = CreateFlatButton("Add Tag to Policy", 480, startY + spacing * 4 - 6, 180, PrimaryBlue);
            btnAdd.Click += BtnAdd_Click;
            cardCreate.Controls.Add(btnAdd);
            mainContent.Controls.Add(cardCreate);

            // Tags Overview DataGrid Card
            Panel cardGrid = CreateCard(24, 540, 1300, 420);
            cardGrid.Controls.Add(new Label { Text = "🏷 Sensitivity Tags Overview", Font = new Font("Segoe UI Semibold", 16), ForeColor = TextPrimary, Location = new Point(20, 20), AutoSize = true });
            cardGrid.Controls.Add(new Label { Text = "View, edit, and remove labels. Changes apply once you push the enterprise policy.", Font = new Font("Segoe UI", 10), ForeColor = TextSecondary, Location = new Point(22, 55), AutoSize = true });

            dgvTags = new DataGridView
            {
                Location = new Point(20, 95),
                Size = new Size(1260, 250),
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                EnableHeadersVisualStyles = false,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };

            dgvTags.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgvTags.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F1F5F9");
            dgvTags.ColumnHeadersDefaultCellStyle.ForeColor = ColorTranslator.FromHtml("#475569");
            dgvTags.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 11F);
            dgvTags.ColumnHeadersHeight = 45;
            dgvTags.RowTemplate.Height = 40;
            dgvTags.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#DBEAFE");
            dgvTags.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvTags.GridColor = BorderSoft;
            dgvTags.DefaultCellStyle.Font = new Font("Segoe UI", 11);

            dgvTags.Columns.Add("Platform", "Target Platform");
            dgvTags.Columns.Add("Name", "Tag Name");
            dgvTags.Columns.Add("Text", "Watermark Text");
            dgvTags.Columns.Add("Placement", "Placement");
            dgvTags.Columns.Add("Size", "Size");
            dgvTags.Columns.Add("Color", "Hex Color");
            cardGrid.Controls.Add(dgvTags);

            Button btnRemove = CreateFlatButton("Remove Selected Tag", 20, 360, 220, DangerRed);
            btnRemove.Click += BtnRemove_Click;
            cardGrid.Controls.Add(btnRemove);

            Button btnSave = CreateFlatButton("Push Enterprise Policy", 1040, 360, 240, SuccessGreen);
            btnSave.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnSave.Click += BtnSave_Click;
            cardGrid.Controls.Add(btnSave);

            mainContent.Controls.Add(cardGrid);
        }

        // --- UI HELPER METHODS ---
        private Panel CreateCard(int x, int y, int w, int h)
        {
            Panel p = new Panel { Location = new Point(x, y), Size = new Size(w, h), BackColor = CardBg };
            p.Paint += (s, e) => { ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, BorderSoft, ButtonBorderStyle.Solid); };
            return p;
        }

        private Label CreateFieldLabel(string text, int x, int y)
        {
            return new Label { Text = text, Location = new Point(x, y), Font = new Font("Segoe UI Semibold", 11), ForeColor = TextPrimary, AutoSize = true };
        }

        private TextBox CreateTextBox(int x, int y, int width)
        {
            return new TextBox { Location = new Point(x, y), Width = width, Font = new Font("Segoe UI", 11), BorderStyle = BorderStyle.FixedSingle };
        }

        private ComboBox CreateComboBox(int x, int y, int width, string[] items)
        {
            var cb = new ComboBox { Location = new Point(x, y), Width = width, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 11) };
            if (items.Length > 0) { cb.Items.AddRange(items); cb.SelectedIndex = 0; }
            return cb;
        }

        private Label CreateIconLabel(string icon, int y)
        {
            return new Label { Text = icon, Font = new Font("Segoe UI", 18), ForeColor = Color.White, Location = new Point(20, y), AutoSize = true, Cursor = Cursors.Hand };
        }

        private Button CreateFlatButton(string text, int x, int y, int width, Color bg)
        {
            Button btn = new Button { Text = text, Location = new Point(x, y), Width = width, Height = 40, BackColor = bg, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Segoe UI Semibold", 11) };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // --- LOGIC METHODS ---
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

            dgvTags.ClearSelection();
        }

        private void BtnPickColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog() == DialogResult.OK)
                {
                    txtColorHex.Text = ColorTranslator.ToHtml(Color.FromArgb(cd.Color.ToArgb()));
                    colorPreview.BackColor = cd.Color; // Update the preview box
                }
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
                MessageBox.Show("Policy saved successfully to all endpoints!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Error: {ex.Message}"); }
        }
    }
}