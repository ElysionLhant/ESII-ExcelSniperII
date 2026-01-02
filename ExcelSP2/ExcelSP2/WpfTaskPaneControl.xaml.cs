using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Web.Script.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelSP2
{
    public partial class WpfTaskPaneControl : UserControl
    {
        // State
        private string capturedAddress;
        private string capturedImageBase64;
        private HeaderInfo cachedHeaderInfo;
        private string cachedColumnRange;
        private ObservableCollection<AttachmentItem> attachments = new ObservableCollection<AttachmentItem>();
        private List<PromptPreset> promptPresets;
        private List<MacroPreset> macroPresets;
        private string promptsFilePath;
        private string macrosFilePath;
        private string settingsFilePath;
        private AppSettings currentSettings;

        public WpfTaskPaneControl()
        {
            InitializeComponent();
            itemsAttachments.ItemsSource = attachments;
            InitializePrompts();
            InitializeMacros();
            InitializeSettings();
        }

        // --- Initialization ---

        private void InitializePrompts()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            promptsFilePath = Path.Combine(folder, "prompts.json");

            LoadPrompts();
        }

        private void LoadPrompts()
        {
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
            RefreshPromptCombo();
        }

        private void InitializeMacros()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            macrosFilePath = Path.Combine(folder, "macros.json");

            LoadMacros();
        }

        private void LoadMacros()
        {
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
            RefreshMacroCombo();
        }

        private void InitializeSettings()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ExcelAIPlugin");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            settingsFilePath = Path.Combine(folder, "settings.json");

            LoadSettings();
        }

        private void LoadSettings()
        {
            if (File.Exists(settingsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(settingsFilePath);
                    var serializer = new JavaScriptSerializer();
                    currentSettings = serializer.Deserialize<AppSettings>(json);
                }
                catch { }
            }
            
            if (currentSettings == null)
            {
                currentSettings = new AppSettings 
                { 
                    ApiUrl = "https://api.openai.com/v1", 
                    Model = "gpt-4o" 
                };
            }
        }

        // --- Event Handlers ---

        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (range == null) return;

                capturedAddress = range.Address[false, false];
                lblSelectionInfo.Text = $"Selected: {capturedAddress} ({range.Rows.Count}R x {range.Columns.Count}C)";

                string currentColKey = GetColumnRangeKey(range);
                if (cachedHeaderInfo != null && cachedColumnRange == currentColKey)
                {
                    lblSelectionInfo.Text += " [Header Cached]";
                    btnResetHeader.Visibility = Visibility.Visible;
                }
                else
                {
                    btnResetHeader.Visibility = Visibility.Collapsed;
                }

                range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    var img = System.Windows.Forms.Clipboard.GetImage();
                    if (img != null)
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            byte[] imageBytes = ms.ToArray();
                            capturedImageBase64 = Convert.ToBase64String(imageBytes);
                            
                            // Update Preview
                            BitmapImage bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.StreamSource = new MemoryStream(imageBytes);
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            previewPopupImage.Source = bitmap;
                        }
                        
                        btnPreview.Visibility = Visibility.Visible;
                        btnClearCapture.Visibility = Visibility.Visible;
                        lblSelectionInfo.Text += " [Image Captured]";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error capturing: " + ex.Message);
            }
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (previewPopupImage.Source == null) return;
            
            // Save to temp and open
            string tempPath = Path.Combine(Path.GetTempPath(), "excel_plugin_preview.png");
            using (var fileStream = new FileStream(tempPath, FileMode.Create))
            {
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create((BitmapSource)previewPopupImage.Source));
                encoder.Save(fileStream);
            }
            System.Diagnostics.Process.Start(tempPath);
        }

        private void BtnPreview_MouseEnter(object sender, MouseEventArgs e)
        {
            if (previewPopupImage.Source != null)
            {
                previewPopup.IsOpen = true;
            }
        }

        private void BtnPreview_MouseLeave(object sender, MouseEventArgs e)
        {
            previewPopup.IsOpen = false;
        }

        private void BtnClearCapture_Click(object sender, RoutedEventArgs e)
        {
            capturedImageBase64 = null;
            capturedAddress = null;
            previewPopupImage.Source = null;
            
            btnPreview.Visibility = Visibility.Collapsed;
            btnClearCapture.Visibility = Visibility.Collapsed;
            lblSelectionInfo.Text = "No selection captured.";
            previewPopup.IsOpen = false;
        }

        private void BtnResetHeader_Click(object sender, RoutedEventArgs e)
        {
            cachedHeaderInfo = null;
            cachedColumnRange = null;
            btnResetHeader.Visibility = Visibility.Collapsed;
            if (lblSelectionInfo.Text.Contains(" [Header Cached]"))
            {
                lblSelectionInfo.Text = lblSelectionInfo.Text.Replace(" [Header Cached]", "");
            }
            MessageBox.Show("Header cache cleared.");
        }

        private void BtnAddFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.Title = "Select Context Files";

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    AddAttachment(file);
                }
            }
        }

        private void BtnRemoveAttachment_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null && btn.Tag is string filePath)
            {
                var item = attachments.FirstOrDefault(a => a.FilePath == filePath);
                if (item != null)
                {
                    attachments.Remove(item);
                }
            }
        }

        private void TxtContext_KeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.V)
            {
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    try
                    {
                        var img = System.Windows.Forms.Clipboard.GetImage();
                        string tempPath = Path.Combine(Path.GetTempPath(), $"pasted_image_{DateTime.Now.Ticks}.png");
                        img.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);
                        AddAttachment(tempPath);
                        e.Handled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to paste image: " + ex.Message);
                    }
                }
                else if (System.Windows.Forms.Clipboard.ContainsFileDropList())
                {
                    var files = System.Windows.Forms.Clipboard.GetFileDropList();
                    foreach (string file in files)
                    {
                        AddAttachment(file);
                    }
                    e.Handled = true;
                }
            }
        }

        private void Context_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effects = DragDropEffects.Copy;
        }

        private void Context_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    AddAttachment(file);
                }
            }
        }

        private void AddAttachment(string filePath)
        {
            if (attachments.Any(a => a.FilePath == filePath)) return;
            attachments.Add(new AttachmentItem { FilePath = filePath, FileName = Path.GetFileName(filePath) });
        }

        // --- Prompt Logic ---

        private void RefreshPromptCombo()
        {
            cmbPrompts.ItemsSource = null;
            cmbPrompts.ItemsSource = promptPresets;
            if (promptPresets.Count > 0) cmbPrompts.SelectedIndex = 0;
        }

        private void CmbPrompts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPrompts.SelectedItem is PromptPreset preset)
            {
                txtPrompt.Text = preset.Content;
            }
        }

        private void BtnSavePrompt_Click(object sender, RoutedEventArgs e)
        {
            string title = Microsoft.VisualBasic.Interaction.InputBox("Enter name for this prompt preset:", "Save Prompt", "New Prompt");
            if (!string.IsNullOrWhiteSpace(title))
            {
                promptPresets.Add(new PromptPreset { Title = title, Content = txtPrompt.Text });
                SavePrompts();
                RefreshPromptCombo();
                cmbPrompts.SelectedIndex = promptPresets.Count - 1;
            }
        }

        private void BtnDeletePrompt_Click(object sender, RoutedEventArgs e)
        {
            if (cmbPrompts.SelectedIndex >= 0)
            {
                if (MessageBox.Show("Delete this preset?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    promptPresets.RemoveAt(cmbPrompts.SelectedIndex);
                    SavePrompts();
                    RefreshPromptCombo();
                    txtPrompt.Text = "";
                }
            }
        }

        private void SavePrompts()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(promptPresets);
            File.WriteAllText(promptsFilePath, json);
        }

        // --- Macro Logic ---

        private void RefreshMacroCombo()
        {
            cmbMacros.ItemsSource = null;
            cmbMacros.ItemsSource = macroPresets;
            if (macroPresets.Count > 0) cmbMacros.SelectedIndex = 0;
        }

        private void CmbMacros_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbMacros.SelectedItem is MacroPreset preset)
            {
                txtMacroCode.Text = preset.Code;
            }
        }

        private void BtnSaveMacro_Click(object sender, RoutedEventArgs e)
        {
            string title = Microsoft.VisualBasic.Interaction.InputBox("Enter name for this macro:", "Save Macro", "New Macro");
            if (!string.IsNullOrWhiteSpace(title))
            {
                macroPresets.Add(new MacroPreset { Title = title, Code = txtMacroCode.Text });
                SaveMacros();
                RefreshMacroCombo();
                cmbMacros.SelectedIndex = macroPresets.Count - 1;
            }
        }

        private void BtnDeleteMacro_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMacros.SelectedIndex >= 0)
            {
                if (MessageBox.Show("Delete this macro?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    macroPresets.RemoveAt(cmbMacros.SelectedIndex);
                    SaveMacros();
                    RefreshMacroCombo();
                    txtMacroCode.Text = "";
                }
            }
        }

        private void SaveMacros()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(macroPresets);
            File.WriteAllText(macrosFilePath, json);
        }

        private void BtnRunMacro_Click(object sender, RoutedEventArgs e)
        {
            string code = txtMacroCode.Text;
            if (string.IsNullOrWhiteSpace(code)) return;

            try
            {
                var match = Regex.Match(code, @"Sub\s+(\w+)", RegexOptions.IgnoreCase);
                if (!match.Success)
                {
                    MessageBox.Show("Could not find a 'Sub Name()' in the code. Please ensure your macro starts with 'Sub Name()'.");
                    return;
                }
                string macroName = match.Groups[1].Value;

                Excel.Application app = Globals.ThisAddIn.Application;
                dynamic vbProj = app.VBE.ActiveVBProject;
                dynamic vbComp = vbProj.VBComponents.Add(1); 
                
                try 
                {
                    vbComp.CodeModule.AddFromString(code);
                    app.Run(macroName);
                }
                finally
                {
                    vbProj.VBComponents.Remove(vbComp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error running macro. \n\nIMPORTANT: You must enable 'Trust access to the VBA project object model' in Excel Options -> Trust Center -> Trust Center Settings -> Macro Settings.\n\nDetails: " + ex.Message);
            }
        }

        // --- Settings Logic ---

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow(promptsFilePath, macrosFilePath, settingsFilePath);
            settingsWindow.SettingsSaved += () =>
            {
                LoadSettings();
                LoadPrompts();
                LoadMacros();
            };
            settingsWindow.ShowDialog();
        }

        private LLMConfig GetConfig(string type)
        {
            if (currentSettings == null) LoadSettings();

            if (!currentSettings.IsAdvancedMode)
            {
                return new LLMConfig 
                { 
                    Provider = currentSettings.Model?.Contains("gpt") == true ? "OpenAI" : "Ollama", 
                    ApiUrl = currentSettings.ApiUrl, 
                    ApiKey = currentSettings.ApiKey, 
                    Model = currentSettings.Model 
                };
            }

            switch (type)
            {
                case "Header": return currentSettings.HeaderDetectionLLM;
                case "Write": return currentSettings.DataWriteLLM;
                case "Op": return currentSettings.DataOpLLM;
                case "Vba": return currentSettings.VBASelfHealingLLM;
                default: return null;
            }
        }

        private void ToggleMode_Checked(object sender, RoutedEventArgs e)
        {
            // Data Op Mode
        }

        private void ToggleMode_Unchecked(object sender, RoutedEventArgs e)
        {
            // Data Write Mode
        }

        // --- Main Execution Logic ---

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            string prompt = txtPrompt.Text;
            string manualContext = txtContext.Text;

            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            if (toggleMode.IsChecked == true)
            {
                // Data Op Mode
                var config = GetConfig("Op");
                if (string.IsNullOrEmpty(config.ApiKey) && config.Provider == "OpenAI")
                {
                    MessageBox.Show("Please set API Key in Settings.");
                    return;
                }

                btnRun.IsEnabled = false;
                try
                {
                    await RunDataMode(config, prompt, manualContext);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    btnRun.IsEnabled = true;
                    lblStatus.Text = "Ready";
                }
                return;
            }

            // Data Write Mode
            var writeConfig = GetConfig("Write");
            if (string.IsNullOrEmpty(writeConfig.ApiKey) && writeConfig.Provider == "OpenAI")
            {
                MessageBox.Show("Please set API Key in Settings.");
                return;
            }

            if (string.IsNullOrEmpty(capturedAddress) || string.IsNullOrEmpty(capturedImageBase64))
            {
                MessageBox.Show("Please capture a selection first.");
                return;
            }

            btnRun.IsEnabled = false;
            lblStatus.Text = "Processing...";

            try
            {
                bool isNewHeaderDetection = false;

                if (cachedHeaderInfo == null)
                {
                    var headerConfig = GetConfig("Header");
                    isNewHeaderDetection = true;
                    lblStatus.Text = "Detecting Header...";
                    cachedHeaderInfo = await DetectHeader(capturedImageBase64, headerConfig);
                    
                    Excel.Range rangeForCache = Globals.ThisAddIn.Application.Range[capturedAddress];
                    cachedColumnRange = GetColumnRangeKey(rangeForCache);
                    
                    if (!lblSelectionInfo.Text.Contains(" [Header Cached]"))
                    {
                        lblSelectionInfo.Text += " [Header Cached]";
                        btnResetHeader.Visibility = Visibility.Visible;
                    }
                }

                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range originalRange = sheet.Range[capturedAddress];
                
                int startRow = originalRange.Row;
                
                if (isNewHeaderDetection)
                {
                    startRow += cachedHeaderInfo.HeaderRows;
                }

                int endRow = originalRange.Row + originalRange.Rows.Count - 1;
                int startCol = originalRange.Column;
                int endCol = originalRange.Column + originalRange.Columns.Count - 1;

                if (startRow > endRow)
                {
                     startRow = endRow + 1; 
                     endRow = startRow; 
                }

                Excel.Range writeRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];

                writeRange.ClearContents();

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

                foreach (var att in attachments)
                {
                    string file = att.FilePath;
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

                var userContent = new List<object>();
                userContent.Add(new { type = "text", text = prompt + "\n\n" + contextBuilder.ToString() });

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
                    model = writeConfig.Model,
                    messages = messages,
                    max_tokens = 16384,
                    temperature = 0.1
                };

                lblStatus.Text = "Generating Content...";

                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(2);
                    if (!string.IsNullOrEmpty(writeConfig.ApiKey))
                        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {writeConfig.ApiKey}");

                    var serializer = new JavaScriptSerializer();
                    string json = serializer.Serialize(requestBody);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync($"{writeConfig.ApiUrl}/chat/completions", content);
                    string responseString = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"API Error: {response.StatusCode}\n{responseString}");

                    dynamic result = serializer.Deserialize<dynamic>(responseString);
                    string llmContent = result["choices"][0]["message"]["content"];
                    
                    llmContent = llmContent.Replace("```json", "").Replace("```", "").Trim();
                    
                    var rows = serializer.Deserialize<dynamic>(llmContent);

                    lblStatus.Text = "Writing to Excel...";
                    WriteToExcelWithDynamicRows(rows, writeRange);
                    lblStatus.Text = "Done!";
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                btnRun.IsEnabled = true;
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

            int availableRows = targetRange.Rows.Count;
            if (dataRowCount > availableRows)
            {
                int rowsToAdd = dataRowCount - availableRows;
                Excel.Range lastRow = targetRange.Rows[targetRange.Rows.Count];
                Excel.Range insertRange = lastRow.Resize[rowsToAdd, targetRange.Columns.Count];
                insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                
                Excel.Worksheet sheet = targetRange.Worksheet;
                targetRange = sheet.Range[targetRange.Cells[1, 1], targetRange.Cells[dataRowCount, targetRange.Columns.Count]];
            }

            Excel.Range finalWriteRange = targetRange.Cells[1, 1].Resize[dataRowCount, dataColCount];
            finalWriteRange.Value2 = data;
            finalWriteRange.Select();
        }

        private string GetColumnRangeKey(Excel.Range range)
        {
            int startCol = range.Column;
            int endCol = range.Column + range.Columns.Count - 1;
            return $"{startCol}-{endCol}";
        }

        private async Task<HeaderInfo> DetectHeader(string imageBase64, LLMConfig config)
        {
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
                model = config.Model,
                messages = messages,
                max_tokens = 4096,
                temperature = 0.0,
                response_format = new { type = "json_object" }
            };

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(1);
                if (!string.IsNullOrEmpty(config.ApiKey))
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {config.ApiKey}");

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{config.ApiUrl}/chat/completions", content);
                string responseString = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                    throw new Exception($"Header Detection Failed: {response.StatusCode}\n{responseString}");

                dynamic result = serializer.Deserialize<dynamic>(responseString);
                string llmContent = result["choices"][0]["message"]["content"];
                
                return serializer.Deserialize<HeaderInfo>(llmContent);
            }
        }

        private async Task RunDataMode(LLMConfig config, string prompt, string manualContext)
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (range == null)
            {
                MessageBox.Show("Please select a range first.");
                return;
            }

            string csvData = GetRangeCsv(range);
            
            StringBuilder contextBuilder = new StringBuilder();
            contextBuilder.AppendLine("--- Selected Data (CSV) ---");
            contextBuilder.AppendLine(csvData);
            
            if (!string.IsNullOrWhiteSpace(manualContext))
            {
                contextBuilder.AppendLine("--- Manual Context ---");
                contextBuilder.AppendLine(manualContext);
            }

            foreach (var att in attachments)
            {
                string file = att.FilePath;
                if (File.Exists(file))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".txt" || ext == ".csv" || ext == ".json" || ext == ".md")
                    {
                        contextBuilder.AppendLine($"--- File: {Path.GetFileName(file)} ---");
                        contextBuilder.AppendLine(File.ReadAllText(file));
                    }
                }
            }

            var messages = new List<object>
            {
                new { role = "system", content = "You are an expert Excel VBA developer. Your task is to write a VBA Sub to perform the user's requested action on the selected data. \n" +
                                                 "RULES:\n" +
                                                 "1. The code MUST be a valid VBA Sub named 'AI_Generated_Action'.\n" +
                                                 "2. The code should operate on the currently selected range (Selection) or the active sheet as appropriate.\n" +
                                                 "3. Return ONLY the VBA code. Do not include markdown formatting like ```vba ... ```. Just the code.\n" +
                                                 "4. Do not use MsgBox unless explicitly asked." },
                new { role = "user", content = prompt + "\n\n" + contextBuilder.ToString() }
            };

            var requestBody = new
            {
                model = config.Model,
                messages = messages,
                max_tokens = 4096,
                temperature = 0.1
            };

            lblStatus.Text = "Generating VBA...";

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(2);
                if (!string.IsNullOrEmpty(config.ApiKey))
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {config.ApiKey}");

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{config.ApiUrl}/chat/completions", content);
                string responseString = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                    throw new Exception($"API Error: {response.StatusCode}\n{responseString}");

                dynamic result = serializer.Deserialize<dynamic>(responseString);
                string vbaCode = result["choices"][0]["message"]["content"];
                
                vbaCode = vbaCode.Replace("```vba", "").Replace("```", "").Trim();

                lblStatus.Text = "Running Macro...";
                RunGeneratedMacro(vbaCode);
                lblStatus.Text = "Done (Data Mode)";
            }
        }

        private string GetRangeCsv(Excel.Range range)
        {
            StringBuilder sb = new StringBuilder();
            object[,] values = range.Value2 as object[,];
            
            if (values == null) return "";

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    object val = values[i, j];
                    string strVal = val != null ? val.ToString() : "";
                    if (strVal.Contains(",") || strVal.Contains("\"") || strVal.Contains("\n"))
                    {
                        strVal = "\"" + strVal.Replace("\"", "\"\"") + "\"";
                    }
                    sb.Append(strVal);
                    if (j < cols) sb.Append(",");
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private void RunGeneratedMacro(string code)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                dynamic vbProj = app.VBE.ActiveVBProject;
                
                dynamic targetComponent = null;
                foreach (dynamic comp in vbProj.VBComponents)
                {
                    if (comp.Name == "Sniper_Temp_Runner")
                    {
                        targetComponent = comp;
                        break;
                    }
                }

                if (targetComponent != null)
                {
                    vbProj.VBComponents.Remove(targetComponent);
                }

                targetComponent = vbProj.VBComponents.Add(1); 
                targetComponent.Name = "Sniper_Temp_Runner";

                targetComponent.CodeModule.AddFromString(code);

                app.Run("Sniper_Temp_Runner.AI_Generated_Action");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error executing macro: " + ex.Message + "\n\nEnsure 'Trust access to the VBA project object model' is enabled.");
            }
        }
    }

    public class AttachmentItem
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
    }
}
