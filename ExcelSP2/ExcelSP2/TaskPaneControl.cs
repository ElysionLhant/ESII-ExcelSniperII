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

    public partial class TaskPaneControl : UserControl
    {
        // UI Controls
        private Button btnCapture;
        private PictureBox picPreview;
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
            
            btnCapture = new Button { Text = "Capture Selection", Width = width, Height = 30, BackColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(3, 3, 3, 3) };
            btnCapture.Click += BtnCapture_Click;
            panel.Controls.Add(btnCapture);

            lblSelectionInfo = new Label { Text = "No selection captured.", Width = width, ForeColor = Color.Gray, AutoSize = true, Margin = new Padding(3, 5, 3, 5) };
            panel.Controls.Add(lblSelectionInfo);

            picPreview = new PictureBox { Width = width, Height = 100, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, Visible = false, Margin = new Padding(3, 5, 3, 5) };
            panel.Controls.Add(picPreview);

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

                // Capture Image
                range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                if (Clipboard.ContainsImage())
                {
                    Image img = Clipboard.GetImage();
                    picPreview.Image = img;
                    picPreview.Visible = true;
                    
                    // Adjust layout if needed, but for now we just show it in the reserved space or let it overlay?
                    // Since we didn't reserve space in y, let's just show it. 
                    // In a real app, we'd use FlowLayoutPanel.
                    // For this fix, let's just make sure it doesn't cover other things.
                    // Actually, let's put it in a popup or just leave it hidden for now to avoid layout issues, 
                    // OR insert it dynamically.
                    // Let's just show a message "Image Captured" to keep UI clean.
                    lblSelectionInfo.Text += " [Image Captured]";
                    
                    using (MemoryStream ms = new MemoryStream())
                    {
                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
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
            if (string.IsNullOrEmpty(txtApiKey.Text))
            {
                MessageBox.Show("Please set API Key in Settings.");
                pnlSettings.Visible = true;
                return;
            }

            btnRun.Enabled = false;
            lblStatus.Text = "Reading files...";

            try
            {
                // Fix SSL/TLS error: Enable TLS 1.2 (Required for OpenAI/Modern APIs)
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                // 1. Prepare Context
                StringBuilder contextBuilder = new StringBuilder();
                List<string> additionalImages = new List<string>();

                if (!string.IsNullOrWhiteSpace(txtContext.Text))
                {
                    contextBuilder.AppendLine("--- Manual Context ---");
                    contextBuilder.AppendLine(txtContext.Text);
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

                // 2. Build Payload
                var userContent = new List<object>();
                userContent.Add(new { type = "text", text = txtPrompt.Text + "\n\n" + contextBuilder.ToString() });

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
                    new { role = "system", content = "You are an Excel assistant. Return ONLY a JSON 2D array." },
                    new { role = "user", content = userContent }
                };

                var requestBody = new
                {
                    model = txtModel.Text,
                    messages = messages,
                    max_tokens = 2000,
                    temperature = 0.1
                };

                lblStatus.Text = "Sending to LLM...";

                // 3. Call API
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(2);
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {txtApiKey.Text}");

                    var serializer = new JavaScriptSerializer();
                    string json = serializer.Serialize(requestBody);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync($"{txtApiUrl.Text}/chat/completions", content);
                    string responseString = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"API Error: {response.StatusCode}\n{responseString}");

                    // 4. Parse Response
                    dynamic result = serializer.Deserialize<dynamic>(responseString);
                    string llmContent = result["choices"][0]["message"]["content"];
                    
                    // Clean markdown
                    llmContent = llmContent.Replace("```json", "").Replace("```", "").Trim();
                    
                    var rows = serializer.Deserialize<dynamic>(llmContent);

                    // 5. Write to Excel
                    this.Invoke(new Action(() => {
                        lblStatus.Text = "Writing to Excel...";
                        WriteToExcel(rows);
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

        private void WriteToExcel(dynamic rows)
        {
            int rowCount = 0;
            int colCount = 0;

            if (rows is Array) rowCount = ((Array)rows).Length;
            else if (rows is System.Collections.IList) rowCount = ((System.Collections.IList)rows).Count;

            if (rowCount == 0) return;

            var firstRow = (rows is Array) ? ((Array)rows).GetValue(0) : ((System.Collections.IList)rows)[0];
            if (firstRow is Array) colCount = ((Array)firstRow).Length;
            else if (firstRow is System.Collections.IList) colCount = ((System.Collections.IList)firstRow).Count;

            object[,] data = new object[rowCount, colCount];
            for (int i = 0; i < rowCount; i++)
            {
                var r = (rows is Array) ? ((Array)rows).GetValue(i) : ((System.Collections.IList)rows)[i];
                for (int j = 0; j < colCount; j++)
                {
                    var val = (r is Array) ? ((Array)r).GetValue(j) : ((System.Collections.IList)r)[j];
                    data[i, j] = val;
                }
            }

            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range startRange;

            if (string.IsNullOrEmpty(capturedAddress))
            {
                startRange = Globals.ThisAddIn.Application.ActiveCell;
            }
            else
            {
                try
                {
                    startRange = sheet.Range[capturedAddress].Cells[1, 1];
                }
                catch
                {
                    startRange = Globals.ThisAddIn.Application.ActiveCell;
                }
            }

            Excel.Range targetRange = startRange.Resize[rowCount, colCount];

            targetRange.Value2 = data;
            targetRange.Select();
        }
    }
}
