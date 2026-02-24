using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using EnterpriseDocClassifier.Models; // This connects to your Shared Class Library

namespace DocClassifier.AdminUI
{
    public partial class Form1 : Form
    {
        // Define UI Components
        private CheckBox chkEnforce;
        private TextBox txtName, txtMarkerText, txtColorHex;
        private ComboBox cbPlacement;
        private NumericUpDown numFontSize;
        private DataGridView dgvTags;
        private PluginConfiguration _currentConfig;

        // The secure enterprise path where the config lives
        private readonly string _configDirectory = @"C:\ProgramData\YourCompany\DocClassifier";
        private readonly string _configPath = @"C:\ProgramData\YourCompany\DocClassifier\config.json";

        public Form1()
        {
            InitializeComponent();
            BuildProfessionalUI();
            LoadExistingConfig();
        }

        private void BuildProfessionalUI()
        {
            // 1. Form Setup
            this.Text = "Enterprise Document Classifier - IT Admin Console";
            this.Size = new Size(800, 600);
            this.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            this.BackColor = Color.WhiteSmoke;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 2. Title
            this.Controls.Add(new Label { Text = "Document Sensitivity Settings", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true });

            // 3. Master Policy Checkbox
            chkEnforce = new CheckBox { Text = "Enforce Classification (Block users from saving without a tag)", Location = new Point(25, 60), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.DarkRed };
            this.Controls.Add(chkEnforce);

            // 4. Input Group Box (The Form)
            GroupBox grpInput = new GroupBox { Text = "Add New Sensitivity Tag", Location = new Point(25, 100), Size = new Size(730, 120) };
            this.Controls.Add(grpInput);

            // --- Row 1 of Inputs ---
            grpInput.Controls.Add(new Label { Text = "Tag Name (e.g. Public):", Location = new Point(15, 30), AutoSize = true });
            txtName = new TextBox { Location = new Point(15, 50), Width = 150 };
            grpInput.Controls.Add(txtName);

            grpInput.Controls.Add(new Label { Text = "Watermark Text:", Location = new Point(180, 30), AutoSize = true });
            txtMarkerText = new TextBox { Location = new Point(180, 50), Width = 200 };
            grpInput.Controls.Add(txtMarkerText);

            grpInput.Controls.Add(new Label { Text = "Placement:", Location = new Point(400, 30), AutoSize = true });
            cbPlacement = new ComboBox { Location = new Point(400, 50), Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cbPlacement.Items.AddRange(new string[] { "Header", "Footer" });
            cbPlacement.SelectedIndex = 0;
            grpInput.Controls.Add(cbPlacement);

            grpInput.Controls.Add(new Label { Text = "Size:", Location = new Point(515, 30), AutoSize = true });
            numFontSize = new NumericUpDown { Location = new Point(515, 50), Width = 50, Minimum = 8, Maximum = 72, Value = 12 };
            grpInput.Controls.Add(numFontSize);

            grpInput.Controls.Add(new Label { Text = "Color (Hex):", Location = new Point(580, 30), AutoSize = true });
            txtColorHex = new TextBox { Location = new Point(580, 50), Width = 70, Text = "#FF0000", ReadOnly = true };
            grpInput.Controls.Add(txtColorHex);

            Button btnPickColor = new Button { Text = "...", Location = new Point(655, 49), Width = 30 };
            btnPickColor.Click += BtnPickColor_Click;
            grpInput.Controls.Add(btnPickColor);

            // --- Row 2 of Inputs ---
            Button btnAdd = new Button { Text = "Add Tag", Location = new Point(15, 85), Width = 100, BackColor = Color.LightSteelBlue };
            btnAdd.Click += BtnAdd_Click;
            grpInput.Controls.Add(btnAdd);

            // 5. Data Grid View (The Table)
            dgvTags = new DataGridView
            {
                Location = new Point(25, 240),
                Size = new Size(730, 230),
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White
            };
            dgvTags.Columns.Add("Name", "Tag Name");
            dgvTags.Columns.Add("Text", "Watermark Text");
            dgvTags.Columns.Add("Placement", "Placement");
            dgvTags.Columns.Add("Size", "Font Size");
            dgvTags.Columns.Add("Color", "Hex Color");
            this.Controls.Add(dgvTags);

            // 6. Footer Buttons
            Button btnRemove = new Button { Text = "Remove Selected Tag", Location = new Point(25, 490), Width = 170 };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            Button btnSave = new Button { Text = "Save Enterprise Policy", Location = new Point(555, 485), Size = new Size(200, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Segoe UI", 11, FontStyle.Bold) };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);
        }

        private void LoadExistingConfig()
        {
            if (File.Exists(_configPath))
            {
                try
                {
                    string json = File.ReadAllText(_configPath);
                    _currentConfig = JsonSerializer.Deserialize<PluginConfiguration>(json) ?? new PluginConfiguration();
                }
                catch { _currentConfig = new PluginConfiguration(); }
            }
            else
            {
                _currentConfig = new PluginConfiguration();
            }

            if (_currentConfig.Classifications == null)
                _currentConfig.Classifications = new List<ClassificationLabel>();

            chkEnforce.Checked = _currentConfig.EnforceClassification;
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            dgvTags.Rows.Clear();
            foreach (var tag in _currentConfig.Classifications)
            {
                dgvTags.Rows.Add(tag.Name, tag.Marker.Text, tag.Marker.Placement, tag.Marker.FontSize, tag.Marker.FontColor);
            }
        }

        private void BtnPickColor_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                // Convert chosen color to HTML Hex Code
                txtColorHex.Text = ColorTranslator.ToHtml(Color.FromArgb(colorDialog.Color.ToArgb()));
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtMarkerText.Text))
            {
                MessageBox.Show("Please enter a Tag Name and Watermark Text.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var newTag = new ClassificationLabel
            {
                Name = txtName.Text.Trim(),
                Marker = new DocumentMarker
                {
                    Text = txtMarkerText.Text.Trim(),
                    Placement = cbPlacement.Text,
                    FontSize = (int)numFontSize.Value,
                    FontColor = txtColorHex.Text
                }
            };

            _currentConfig.Classifications.Add(newTag);
            RefreshGrid();

            // Clear inputs
            txtName.Clear();
            txtMarkerText.Clear();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (dgvTags.SelectedRows.Count > 0)
            {
                int index = dgvTags.SelectedRows[0].Index;
                _currentConfig.Classifications.RemoveAt(index);
                RefreshGrid();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            _currentConfig.EnforceClassification = chkEnforce.Checked;

            // Make it readable for humans using WriteIndented
            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(_currentConfig, options);

            try
            {
                // Ensure the folder exists before saving
                if (!Directory.Exists(_configDirectory))
                {
                    Directory.CreateDirectory(_configDirectory);
                }

                File.WriteAllText(_configPath, jsonString);
                MessageBox.Show("Enterprise Configuration saved successfully. \n\nYou can now push this file via ManageEngine to update all endpoints.", "Policy Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save configuration: {ex.Message}", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}