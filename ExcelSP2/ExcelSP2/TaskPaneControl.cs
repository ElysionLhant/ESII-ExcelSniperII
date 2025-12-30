using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web.Script.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

using System.Text.RegularExpressions;

namespace ExcelSP2
{
    public class PromptPreset
    {
        public string Title { get; set; }
        public string Content { get; set; }
    }

    public class MacroPreset
    {
        public string Title { get; set; }
        public string Code { get; set; }
    }

    public class AppSettings
    {
        public string ApiUrl { get; set; }
        public string ApiKey { get; set; }
        public string Model { get; set; }
    }

    public class HeaderInfo
    {
        public string HeaderContent { get; set; }
        public int HeaderRows { get; set; }
    }

    public partial class TaskPaneControl : UserControl
    {
        // UI Controls
        private Button btnCapture;
        private Button btnResetHeader; // New Reset Button
        private Button btnPreview; // New Preview Button
        private Image capturedImage; // Store captured image
        private Form previewPopup; // Popup for hover
        private PictureBox previewPopupBox; // PictureBox in popup
        // private PictureBox picPreview; // Removed
        private Label lblSelectionInfo;
        // private Panel pnlDropZone; // Removed
        // private Label lblDropHint; // Removed
        private ListBox lstFiles;
        private TextBox txtContext;
        
        // Prompt Controls
        private ComboBox cmbPrompts;
        private Button btnSavePrompt;
        private Button btnDeletePrompt;
        private TextBox txtPrompt;

        // Macro Controls
        private ComboBox cmbMacros;
        private Button btnRunMacro;
        private Button btnSaveMacro;
        private Button btnDeleteMacro;
        private TextBox txtMacroCode;
        
        private Button btnRun;
        private Button btnSettings;
        private Panel pnlSettings;
        private TextBox txtApiUrl;
        private TextBox txtApiKey;
        private TextBox txtModel;
        private Label lblStatus;

        // State
        private string capturedAddress;
        private string capturedImageBase64;
        private HeaderInfo cachedHeaderInfo; // Cache for header
        private string cachedColumnRange; // Cache for column range
        private List<string> filePaths = new List<string>();
        private List<PromptPreset> promptPresets;
        private List<MacroPreset> macroPresets;
        private string promptsFilePath;
        private string macrosFilePath;
        private string settingsFilePath;

        public TaskPaneControl()
        {
            InitializeComponent();
            InitializePrompts();
            InitializeMacros();
            SetupCustomUI();
            InitializeSettings();
        }

        private void InitializePrompts()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            promptsFilePath = Path.Combine(folder, "prompts.json");

            if (File.Exists(promptsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(promptsFilePath);
                    var serializer = new JavaScriptSerializer();
                    promptPresets = serializer.Deserialize<List<PromptPreset>>(json);
                }
                catch { promptPresets = null; }
            }
            
            if (promptPresets == null || promptPresets.Count == 0)
            {
                promptPresets = new List<PromptPreset>
                {
                    new PromptPreset { Title = "General Fill", Content = "Fill the table based on the provided image and files." },
                    new PromptPreset { Title = "Invoice Extraction", Content = "Extract line items from the invoice image/pdf. Columns: Description, Quantity, Unit Price, Total." },
                    new PromptPreset { Title = "Data Cleanup", Content = "Format the data in the image to be consistent and correct any typos." }
                };
                SavePrompts();
            }
        }

        private void InitializeMacros()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            macrosFilePath = Path.Combine(folder, "macros.json");

            if (File.Exists(macrosFilePath))
            {
                try
                {
                    string json = File.ReadAllText(macrosFilePath);
                    var serializer = new JavaScriptSerializer();
                    macroPresets = serializer.Deserialize<List<MacroPreset>>(json);
                }
                catch { macroPresets = null; }
            }

            if (macroPresets == null)
            {
                macroPresets = new List<MacroPreset>
                {
                    new MacroPreset { Title = "HelloWorld", Code = "Sub HelloWorld()\n    MsgBox \"Hello from VSTO!\"\nEnd Sub" }
                };
                SaveMacros();
            }
        }

        private void SavePrompts()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(promptPresets);
            File.WriteAllText(promptsFilePath, json);
        }

        private void SaveMacros()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(macroPresets);
            File.WriteAllText(macrosFilePath, json);
        }

        private void InitializeSettings()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            settingsFilePath = Path.Combine(folder, "settings.json");

            if (File.Exists(settingsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(settingsFilePath);
                    var serializer = new JavaScriptSerializer();
                    var settings = serializer.Deserialize<AppSettings>(json);
                    if (settings != null)
                    {
                        txtApiUrl.Text = settings.ApiUrl;
                        txtApiKey.Text = settings.ApiKey;
                        txtModel.Text = settings.Model;
                    }
                }
                catch { }
            }
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                ApiUrl = txtApiUrl.Text,
                ApiKey = txtApiKey.Text,
                Model = txtModel.Text
            };
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(settings);
            File.WriteAllText(settingsFilePath, json);
            MessageBox.Show("Settings saved!");
        }

        private void SetupCustomUI()
        {
            this.AutoScroll = true;
            this.Padding = new Padding(10);
            this.BackColor = Color.WhiteSmoke;

            // Use FlowLayoutPanel for automatic layout
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.FlowDirection = FlowDirection.TopDown;
            panel.WrapContents = false;
            panel.AutoScroll = true;
            this.Controls.Add(panel);

            int width = 260; // Slightly less to account for scrollbar

            // 1. Selection Section
            panel.Controls.Add(CreateHeader("1. Selection"));
            
            // Initialize Preview Popup
            previewPopup = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Size = new Size(300, 300),
                TopMost = true,
                BackColor = Color.DimGray,
                Padding = new Padding(1)
            };
            previewPopupBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            previewPopup.Controls.Add(previewPopupBox);

            FlowLayoutPanel pnlCapture = new FlowLayoutPanel { Width = width, Height = 35, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
            
            btnCapture = new Button { Text = "Capture", Width = 110, Height = 30, BackColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(3) };
            btnCapture.Click += BtnCapture_Click;
            pnlCapture.Controls.Add(btnCapture);

            btnPreview = new Button { Text = "Preview", Width = 60, Height = 30, BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat, Margin = new Padding(3), Visible = false };
            btnPreview.Click += BtnPreview_Click;
            btnPreview.MouseEnter += BtnPreview_MouseEnter;
            btnPreview.MouseLeave += BtnPreview_MouseLeave;
            pnlCapture.Controls.Add(btnPreview);

            btnResetHeader = new Button { Text = "Reset", Width = 60, Height = 30, BackColor = Color.LightSalmon, FlatStyle = FlatStyle.Flat, Margin = new Padding(3), Visible = false };
            btnResetHeader.Click += BtnResetHeader_Click;
            pnlCapture.Controls.Add(btnResetHeader);

            panel.Controls.Add(pnlCapture);

            lblSelectionInfo = new Label { Text = "No selection captured.", Width = width, ForeColor = Color.Gray, AutoSize = true, Margin = new Padding(3, 5, 3, 5) };
            panel.Controls.Add(lblSelectionInfo);

            // 2. Context Section
            panel.Controls.Add(CreateHeader("2. Context Materials"));
            
            lstFiles = new ListBox { Width = width, Height = 80, AllowDrop = true, ForeColor = Color.Gray };
            lstFiles.Items.Add("[Drag files here...]");
            lstFiles.DragEnter += LstFiles_DragEnter;
            lstFiles.DragDrop += LstFiles_DragDrop;
            lstFiles.KeyDown += LstFiles_KeyDown;
            panel.Controls.Add(lstFiles);

            Label lblManualContext = new Label { Text = "Or paste text:", Width = width, Height = 15, ForeColor = Color.DimGray, Font = new Font(this.Font.FontFamily, 8), Margin = new Padding(3, 10, 3, 0) };
            panel.Controls.Add(lblManualContext);

            txtContext = new TextBox { Width = width, Height = 60, Multiline = true, ScrollBars = ScrollBars.Vertical };
            panel.Controls.Add(txtContext);

            // 3. Prompt Section
            panel.Controls.Add(CreateHeader("3. Prompt"));

            // Prompt Management Row (Need a sub-panel for horizontal layout)
            FlowLayoutPanel pnlPrompts = new FlowLayoutPanel { Width = width, Height = 30, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
            
            cmbPrompts = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            RefreshPromptCombo();
            cmbPrompts.SelectedIndexChanged += CmbPrompts_SelectedIndexChanged;
            pnlPrompts.Controls.Add(cmbPrompts);

            btnSavePrompt = new Button { Text = "Save", Width = 45, Height = 23, BackColor = Color.White }; 
            btnSavePrompt.Click += BtnSavePrompt_Click;
            pnlPrompts.Controls.Add(btnSavePrompt);

            btnDeletePrompt = new Button { Text = "Del", Width = 40, Height = 23, BackColor = Color.White }; 
            btnDeletePrompt.Click += BtnDeletePrompt_Click;
            pnlPrompts.Controls.Add(btnDeletePrompt);
            
            panel.Controls.Add(pnlPrompts);
            
            txtPrompt = new TextBox { Width = width, Height = 80, Multiline = true, ScrollBars = ScrollBars.Vertical };
            if (promptPresets != null && promptPresets.Count > 0) txtPrompt.Text = promptPresets[0].Content;
            panel.Controls.Add(txtPrompt);

            // 4. Macro Section
            panel.Controls.Add(CreateHeader("4. Macro Library"));

            FlowLayoutPanel pnlMacros = new FlowLayoutPanel { Width = width, Height = 30, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };

            cmbMacros = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            RefreshMacroCombo();
            cmbMacros.SelectedIndexChanged += CmbMacros_SelectedIndexChanged;
            pnlMacros.Controls.Add(cmbMacros);

            btnRunMacro = new Button { Text = "Run", Width = 40, Height = 23, BackColor = Color.LightGreen, FlatStyle = FlatStyle.Flat };
            btnRunMacro.Click += BtnRunMacro_Click;
            pnlMacros.Controls.Add(btnRunMacro);

            btnSaveMacro = new Button { Text = "Save", Width = 40, Height = 23, BackColor = Color.White };
            btnSaveMacro.Click += BtnSaveMacro_Click;
            pnlMacros.Controls.Add(btnSaveMacro);

            btnDeleteMacro = new Button { Text = "Del", Width = 35, Height = 23, BackColor = Color.White };
            btnDeleteMacro.Click += BtnDeleteMacro_Click;
            pnlMacros.Controls.Add(btnDeleteMacro);

            panel.Controls.Add(pnlMacros);

            txtMacroCode = new TextBox { Width = width, Height = 80, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9) };
            if (macroPresets != null && macroPresets.Count > 0) txtMacroCode.Text = macroPresets[0].Code;
            panel.Controls.Add(txtMacroCode);

            // 5. Settings Toggle
            btnSettings = new Button { Text = "⚙️ Settings", Width = 80, Height = 25, Font = new Font(this.Font.FontFamily, 8), Margin = new Padding(3, 10, 3, 3) };
            btnSettings.Click += (s, e) => { pnlSettings.Visible = !pnlSettings.Visible; };
            panel.Controls.Add(btnSettings);

            pnlSettings = new Panel { Width = width, Height = 220, Visible = false, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            SetupSettingsPanel();
            panel.Controls.Add(pnlSettings);

            // 5. Action
            btnRun = new Button { Text = "Generate & Fill", Width = width, Height = 40, BackColor = Color.DodgerBlue, ForeColor = Color.White, Font = new Font(this.Font, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Margin = new Padding(3, 20, 3, 3) };
            btnRun.Click += BtnRun_Click;
            panel.Controls.Add(btnRun);

            lblStatus = new Label { Text = "Ready", Width = width, ForeColor = Color.Blue, AutoSize = true, MaximumSize = new Size(width, 0) };
            panel.Controls.Add(lblStatus);
        }

        private Label CreateHeader(string text)
        {
            return new Label { Text = text, Width = 200, Font = new Font(this.Font, FontStyle.Bold), Margin = new Padding(3, 15, 3, 5) };
        }

        private void SetupSettingsPanel()
        {
            int sy = 5;
            pnlSettings.Controls.Add(new Label { Text = "API URL:", Top = sy, Left = 5 });
            txtApiUrl = new TextBox { Text = "https://api.openai.com/v1", Top = sy + 20, Left = 5, Width = 260 };
            pnlSettings.Controls.Add(txtApiUrl);
            sy += 50;

            pnlSettings.Controls.Add(new Label { Text = "API Key:", Top = sy, Left = 5 });
            txtApiKey = new TextBox { PasswordChar = '*', Top = sy + 20, Left = 5, Width = 260 };
            pnlSettings.Controls.Add(txtApiKey);
            sy += 50;

            pnlSettings.Controls.Add(new Label { Text = "Model:", Top = sy, Left = 5 });
            txtModel = new TextBox { Text = "gpt-4o", Top = sy + 20, Left = 5, Width = 260 };
            pnlSettings.Controls.Add(txtModel);

            sy += 50;
            Button btnSaveSettings = new Button { Text = "Save Settings", Top = sy, Left = 5, Width = 260, Height = 30, BackColor = Color.WhiteSmoke };
            btnSaveSettings.Click += (s, e) => { SaveSettings(); pnlSettings.Visible = false; };
            pnlSettings.Controls.Add(btnSaveSettings);
        }



        private void RefreshPromptCombo()
        {
            cmbPrompts.Items.Clear();
            foreach (var p in promptPresets)
            {
                cmbPrompts.Items.Add(p.Title);
            }
        }

        private void CmbPrompts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPrompts.SelectedIndex >= 0 && cmbPrompts.SelectedIndex < promptPresets.Count)
            {
                txtPrompt.Text = promptPresets[cmbPrompts.SelectedIndex].Content;
            }
        }

        private void BtnSavePrompt_Click(object sender, EventArgs e)
        {
            string title = ShowInputDialog("Enter name for this prompt preset:", "Save Prompt");
            if (!string.IsNullOrWhiteSpace(title))
            {
                promptPresets.Add(new PromptPreset { Title = title, Content = txtPrompt.Text });
                SavePrompts();
                RefreshPromptCombo();
                cmbPrompts.SelectedIndex = promptPresets.Count - 1;
            }
        }

        private void BtnDeletePrompt_Click(object sender, EventArgs e)
        {
            if (cmbPrompts.SelectedIndex >= 0)
            {
                if (MessageBox.Show("Delete this preset?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    promptPresets.RemoveAt(cmbPrompts.SelectedIndex);
                    SavePrompts();
                    RefreshPromptCombo();
                    txtPrompt.Text = "";
                }
            }
        }

        // --- Macro Logic ---

        private void RefreshMacroCombo()
        {
            cmbMacros.Items.Clear();
            foreach (var m in macroPresets)
            {
                cmbMacros.Items.Add(m.Title);
            }
            if (macroPresets.Count > 0) cmbMacros.SelectedIndex = 0;
        }

        private void CmbMacros_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMacros.SelectedIndex >= 0 && cmbMacros.SelectedIndex < macroPresets.Count)
            {
                txtMacroCode.Text = macroPresets[cmbMacros.SelectedIndex].Code;
            }
        }

        private void BtnSaveMacro_Click(object sender, EventArgs e)
        {
            string title = ShowInputDialog("Enter name for this macro:", "Save Macro");
            if (!string.IsNullOrWhiteSpace(title))
            {
                macroPresets.Add(new MacroPreset { Title = title, Code = txtMacroCode.Text });
                SaveMacros();
                RefreshMacroCombo();
                cmbMacros.SelectedIndex = macroPresets.Count - 1;
            }
        }

        private void BtnDeleteMacro_Click(object sender, EventArgs e)
        {
            if (cmbMacros.SelectedIndex >= 0)
            {
                if (MessageBox.Show("Delete this macro?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    macroPresets.RemoveAt(cmbMacros.SelectedIndex);
                    SaveMacros();
                    RefreshMacroCombo();
                    txtMacroCode.Text = "";
                }
            }
        }

        private void BtnRunMacro_Click(object sender, EventArgs e)
        {
            string code = txtMacroCode.Text;
            if (string.IsNullOrWhiteSpace(code)) return;

            try
            {
                // 1. Find the Sub name
                var match = Regex.Match(code, @"Sub\s+(\w+)", RegexOptions.IgnoreCase);
                if (!match.Success)
                {
                    MessageBox.Show("Could not find a 'Sub Name()' in the code. Please ensure your macro starts with 'Sub Name()'.");
                    return;
                }
                string macroName = match.Groups[1].Value;

                // 2. Inject and Run
                Excel.Application app = Globals.ThisAddIn.Application;
                
                // Note: This requires "Trust access to the VBA project object model" in Excel Trust Center
                dynamic vbProj = app.VBE.ActiveVBProject;
                dynamic vbComp = vbProj.VBComponents.Add(1); // 1 = vbext_ct_StdModule
                
                try 
                {
                    vbComp.CodeModule.AddFromString(code);
                    app.Run(macroName);
                }
                finally
                {
                    // Cleanup: Remove the module
                    vbProj.VBComponents.Remove(vbComp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error running macro. \n\nIMPORTANT: You must enable 'Trust access to the VBA project object model' in Excel Options -> Trust Center -> Trust Center Settings -> Macro Settings.\n\nDetails: " + ex.Message);
            }
        }

        // Simple Input Dialog Helper
        private string ShowInputDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 400,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 20, Top = 20, Text = text, Width = 350 };
            TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 340 };
            Button confirmation = new Button() { Text = "Ok", Left = 250, Width = 100, Top = 80, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        // --- Logic ---

        private void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (range == null) return;

                capturedAddress = range.Address[false, false];
                lblSelectionInfo.Text = $"Selected: {capturedAddress} ({range.Rows.Count}R x {range.Columns.Count}C)";

                // Check Cache
                string currentColKey = GetColumnRangeKey(range);
                if (cachedHeaderInfo != null && cachedColumnRange == currentColKey)
                {
                    lblSelectionInfo.Text += " [Header Cached]";
                    btnResetHeader.Visible = true;
                }
                else
                {
                    btnResetHeader.Visible = false;
                }

                // Capture Image
                range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                if (Clipboard.ContainsImage())
                {
                    if (capturedImage != null) capturedImage.Dispose();
                    capturedImage = Clipboard.GetImage();
                    btnPreview.Visible = true;
                    lblSelectionInfo.Text += " [Image Captured]";
                    
                    using (MemoryStream ms = new MemoryStream())
                    {
                        capturedImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        byte[] imageBytes = ms.ToArray();
                        capturedImageBase64 = Convert.ToBase64String(imageBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error capturing: " + ex.Message);
            }
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            if (capturedImage == null) return;
            string tempPath = Path.Combine(Path.GetTempPath(), "excel_plugin_preview.png");
            capturedImage.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);
            System.Diagnostics.Process.Start(tempPath);
        }

        private void BtnPreview_MouseEnter(object sender, EventArgs e)
        {
            if (capturedImage == null) return;
            
            previewPopupBox.Image = capturedImage;
            
            // Calculate position: Left of the button
            Point pt = btnPreview.PointToScreen(Point.Empty);
            previewPopup.Location = new Point(pt.X - previewPopup.Width - 5, pt.Y);
            
            previewPopup.Show();
        }

        private void BtnPreview_MouseLeave(object sender, EventArgs e)
        {
            previewPopup.Hide();
        }

        private void LstFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void LstFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (lstFiles.Items.Count == 1 && lstFiles.Items[0].ToString() == "[Drag files here...]")
            {
                lstFiles.Items.Clear();
                lstFiles.ForeColor = Color.Black;
            }

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (!filePaths.Contains(file))
                {
                    filePaths.Add(file);
                    lstFiles.Items.Add(Path.GetFileName(file));
                }
            }
        }

        private void LstFiles_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) && lstFiles.SelectedIndex >= 0)
            {
                if (lstFiles.Items[lstFiles.SelectedIndex].ToString() == "[Drag files here...]") return;
                
                int index = lstFiles.SelectedIndex;
                filePaths.RemoveAt(index);
                lstFiles.Items.RemoveAt(index);

                if (lstFiles.Items.Count == 0)
                {
                    lstFiles.Items.Add("[Drag files here...]");
                    lstFiles.ForeColor = Color.Gray;
                }
            }
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            // Capture UI values on UI thread
            string apiKey = txtApiKey.Text;
            string apiUrl = txtApiUrl.Text;
            string model = txtModel.Text;
            string prompt = txtPrompt.Text;
            string manualContext = txtContext.Text;

            if (string.IsNullOrEmpty(apiKey))
            {
                MessageBox.Show("Please set API Key in Settings.");
                pnlSettings.Visible = true;
                return;
            }

            if (string.IsNullOrEmpty(capturedAddress) || string.IsNullOrEmpty(capturedImageBase64))
            {
                MessageBox.Show("Please capture a selection first.");
                return;
            }

            btnRun.Enabled = false;
            lblStatus.Text = "Processing...";

            try
            {
                // Fix SSL/TLS error: Enable TLS 1.2 (Required for OpenAI/Modern APIs)
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                bool isNewHeaderDetection = false;

                // Step 1: Header Detection
                if (cachedHeaderInfo == null)
                {
                    isNewHeaderDetection = true;
                    lblStatus.Text = "Detecting Header...";
                    cachedHeaderInfo = await DetectHeader(capturedImageBase64, apiUrl, apiKey, model);
                    
                    // Update Cache Key
                    Excel.Range rangeForCache = Globals.ThisAddIn.Application.Range[capturedAddress];
                    cachedColumnRange = GetColumnRangeKey(rangeForCache);
                    
                    // Update UI to show cached status
                    this.Invoke(new Action(() => {
                        if (!lblSelectionInfo.Text.Contains(" [Header Cached]"))
                        {
                            lblSelectionInfo.Text += " [Header Cached]";
                            btnResetHeader.Visible = true;
                        }
                    }));
                }

                // Step 2: Adjust Selection Range
                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range originalRange = sheet.Range[capturedAddress];
                
                int startRow = originalRange.Row;
                
                // Only apply header offset if this is a fresh detection (Case A: Initial Capture)
                // If using cached header (Case B: Continuous Input), we write to the exact selected area
                if (isNewHeaderDetection)
                {
                    startRow += cachedHeaderInfo.HeaderRows;
                }

                int endRow = originalRange.Row + originalRange.Rows.Count - 1;
                int startCol = originalRange.Column;
                int endCol = originalRange.Column + originalRange.Columns.Count - 1;

                // Ensure we have a valid range
                if (startRow > endRow)
                {
                     // If header takes up all rows, start writing from next row
                     startRow = endRow + 1; 
                     endRow = startRow; // Initial write range is 1 row
                }

                Excel.Range writeRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];

                // Step 3: Prepare Data (Clear Value, Keep Format)
                this.Invoke(new Action(() => {
                    writeRange.ClearContents(); // Clears value/formula but keeps format
                }));

                // Prepare Context for Execution LLM
                StringBuilder contextBuilder = new StringBuilder();
                contextBuilder.AppendLine($"--- Header Information ---");
                contextBuilder.AppendLine($"Header Content: {cachedHeaderInfo.HeaderContent}");
                contextBuilder.AppendLine("(NOTE: This header is provided for context only. Do NOT include it in the output data.)");
                contextBuilder.AppendLine($"Target Write Start Row: {startRow}");
                
                List<string> additionalImages = new List<string>();

                if (!string.IsNullOrWhiteSpace(manualContext))
                {
                    contextBuilder.AppendLine("--- Manual Context ---");
                    contextBuilder.AppendLine(manualContext);
                    contextBuilder.AppendLine();
                }

                foreach (string file in filePaths)
                {
                    if (File.Exists(file))
                    {
                        string ext = Path.GetExtension(file).ToLower();
                        if (ext == ".txt" || ext == ".csv" || ext == ".json" || ext == ".md")
                        {
                            contextBuilder.AppendLine($"--- File: {Path.GetFileName(file)} ---");
                            contextBuilder.AppendLine(File.ReadAllText(file));
                        }
                        else if (ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif")
                        {
                            try
                            {
                                byte[] bytes = File.ReadAllBytes(file);
                                string base64 = Convert.ToBase64String(bytes);
                                additionalImages.Add(base64);
                                contextBuilder.AppendLine($"--- Image File: {Path.GetFileName(file)} (Attached) ---");
                            }
                            catch { }
                        }
                        else
                        {
                            contextBuilder.AppendLine($"--- File: {Path.GetFileName(file)} ---");
                            contextBuilder.AppendLine("[Binary/PDF content reading requires NuGet packages. Filename provided for context.]");
                        }
                    }
                }

                // Build Payload
                var userContent = new List<object>();
                userContent.Add(new { type = "text", text = prompt + "\n\n" + contextBuilder.ToString() });

                // We send the original captured image as reference, plus any additional images
                if (!string.IsNullOrEmpty(capturedImageBase64))
                {
                    userContent.Add(new {
                        type = "image_url",
                        image_url = new { url = $"data:image/png;base64,{capturedImageBase64}" }
                    });
                }

                foreach (var imgBase64 in additionalImages)
                {
                    userContent.Add(new {
                        type = "image_url",
                        image_url = new { url = $"data:image/png;base64,{imgBase64}" }
                    });
                }

                var messages = new List<object>
                {
                    new { role = "system", content = "You are an Excel assistant. Return ONLY a JSON 2D array of values to fill the table. IMPORTANT: Do NOT include the header row in your output. Only return the data rows." },
                    new { role = "user", content = userContent }
                };

                var requestBody = new
                {
                    model = model,
                    messages = messages,
                    max_tokens = 16384,
                    temperature = 0.1
                };

                this.Invoke(new Action(() => lblStatus.Text = "Generating Content..."));

                // Call API
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(2);
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                    var serializer = new JavaScriptSerializer();
                    string json = serializer.Serialize(requestBody);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync($"{apiUrl}/chat/completions", content);
                    string responseString = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"API Error: {response.StatusCode}\n{responseString}");

                    // Parse Response
                    dynamic result = serializer.Deserialize<dynamic>(responseString);
                    string llmContent = result["choices"][0]["message"]["content"];
                    
                    // Clean markdown
                    llmContent = llmContent.Replace("```json", "").Replace("```", "").Trim();
                    
                    var rows = serializer.Deserialize<dynamic>(llmContent);

                    // Step 5 & 6: Write to Excel with Dynamic Rows
                    this.Invoke(new Action(() => {
                        lblStatus.Text = "Writing to Excel...";
                        WriteToExcelWithDynamicRows(rows, writeRange);
                        lblStatus.Text = "Done!";
                    }));
                }
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(() => {
                    lblStatus.Text = "Error: " + ex.Message;
                    MessageBox.Show(ex.ToString());
                }));
            }
            finally
            {
                this.Invoke(new Action(() => btnRun.Enabled = true));
            }
        }

        private void WriteToExcelWithDynamicRows(dynamic rows, Excel.Range targetRange)
        {
            int dataRowCount = 0;
            int dataColCount = 0;

            if (rows is Array) dataRowCount = ((Array)rows).Length;
            else if (rows is System.Collections.IList) dataRowCount = ((System.Collections.IList)rows).Count;

            if (dataRowCount == 0) return;

            var firstRow = (rows is Array) ? ((Array)rows).GetValue(0) : ((System.Collections.IList)rows)[0];
            if (firstRow is Array) dataColCount = ((Array)firstRow).Length;
            else if (firstRow is System.Collections.IList) dataColCount = ((System.Collections.IList)firstRow).Count;

            object[,] data = new object[dataRowCount, dataColCount];
            for (int i = 0; i < dataRowCount; i++)
            {
                var r = (rows is Array) ? ((Array)rows).GetValue(i) : ((System.Collections.IList)rows)[i];
                for (int j = 0; j < dataColCount; j++)
                {
                    var val = (r is Array) ? ((Array)r).GetValue(j) : ((System.Collections.IList)r)[j];
                    data[i, j] = val;
                }
            }

            // Check if we need to insert rows
            int availableRows = targetRange.Rows.Count;
            if (dataRowCount > availableRows)
            {
                int rowsToAdd = dataRowCount - availableRows;
                
                // Insert rows starting from the last row of the target range
                // We use the last row as the anchor to insert below/at
                Excel.Range lastRow = targetRange.Rows[targetRange.Rows.Count];
                
                // Resize to cover the number of rows we need to add
                Excel.Range insertRange = lastRow.Resize[rowsToAdd, targetRange.Columns.Count];
                
                // Insert shifting down, copying format from above (default usually works well for tables)
                insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                
                // Note: After insertion, the targetRange variable might not automatically expand in the way we want
                // So we redefine the target range to start from the same top-left but extend to new height
                Excel.Worksheet sheet = targetRange.Worksheet;
                targetRange = sheet.Range[targetRange.Cells[1, 1], targetRange.Cells[dataRowCount, targetRange.Columns.Count]];
            }

            // Write data
            // We write to the top-left of the target range, resizing to match data dimensions
            Excel.Range finalWriteRange = targetRange.Cells[1, 1].Resize[dataRowCount, dataColCount];
            finalWriteRange.Value2 = data;
            finalWriteRange.Select();
        }

        private void BtnResetHeader_Click(object sender, EventArgs e)
        {
            cachedHeaderInfo = null;
            cachedColumnRange = null;
            btnResetHeader.Visible = false;
            if (lblSelectionInfo.Text.Contains(" [Header Cached]"))
            {
                lblSelectionInfo.Text = lblSelectionInfo.Text.Replace(" [Header Cached]", "");
            }
            MessageBox.Show("Header cache cleared.");
        }

        private string GetColumnRangeKey(Excel.Range range)
        {
            // Returns a key representing the columns, e.g., "1-4" for Columns A to D
            int startCol = range.Column;
            int endCol = range.Column + range.Columns.Count - 1;
            return $"{startCol}-{endCol}";
        }

        private async Task<HeaderInfo> DetectHeader(string imageBase64, string apiUrl, string apiKey, string model)
        {
            // Prompt for header detection
            var messages = new List<object>
            {
                new { role = "system", content = "You are an expert at analyzing Excel tables. Your task is to identify the header rows in the provided image of a table." },
                new { role = "user", content = new List<object> {
                    new { type = "text", text = "Analyze this image. Identify the header content and the number of rows the header occupies. Return a JSON object with keys: 'HeaderContent' (string, the text of the headers) and 'HeaderRows' (integer, the count of header rows). Example: { \"HeaderContent\": \"Name | Age | Date\", \"HeaderRows\": 1 }" },
                    new { type = "image_url", image_url = new { url = $"data:image/png;base64,{imageBase64}" } }
                }}
            };

            var requestBody = new
            {
                model = model, // Use the configured model
                messages = messages,
                max_tokens = 4096,
                temperature = 0.0,
                response_format = new { type = "json_object" } // Force JSON
            };

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(1);
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{apiUrl}/chat/completions", content);
                string responseString = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                    throw new Exception($"Header Detection Failed: {response.StatusCode}\n{responseString}");

                dynamic result = serializer.Deserialize<dynamic>(responseString);
                string llmContent = result["choices"][0]["message"]["content"];
                
                return serializer.Deserialize<HeaderInfo>(llmContent);
            }
        }
    }
}
