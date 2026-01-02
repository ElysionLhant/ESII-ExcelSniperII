using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Web.Script.Serialization;

namespace ExcelSP2
{
    public partial class SettingsWindow : Window
    {
        public List<PromptPreset> Prompts { get; private set; }
        public List<MacroPreset> Macros { get; private set; }
        public List<MacroPreset> InFileMacros { get; private set; }
        public AppSettings Settings { get; private set; }

        private string promptsFilePath;
        private string macrosFilePath;
        private string settingsFilePath;

        public event Action SettingsSaved;

        public SettingsWindow(string promptsPath, string macrosPath, string settingsPath)
        {
            InitializeComponent();
            promptsFilePath = promptsPath;
            macrosFilePath = macrosPath;
            settingsFilePath = settingsPath;

            LoadData();
        }

        private void LoadData()
        {
            // Load Prompts
            if (File.Exists(promptsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(promptsFilePath);
                    var serializer = new JavaScriptSerializer();
                    Prompts = serializer.Deserialize<List<PromptPreset>>(json) ?? new List<PromptPreset>();
                }
                catch { Prompts = new List<PromptPreset>(); }
            }
            else Prompts = new List<PromptPreset>();

            lstPrompts.ItemsSource = null;
            lstPrompts.ItemsSource = Prompts;

            // Load Macros
            if (File.Exists(macrosFilePath))
            {
                try
                {
                    string json = File.ReadAllText(macrosFilePath);
                    var serializer = new JavaScriptSerializer();
                    Macros = serializer.Deserialize<List<MacroPreset>>(json) ?? new List<MacroPreset>();
                }
                catch { Macros = new List<MacroPreset>(); }
            }
            else Macros = new List<MacroPreset>();

            lstMacros.ItemsSource = null;
            lstMacros.ItemsSource = Macros;

            // Load In-File Macros
            LoadInFileMacros();
            lstInFileMacros.ItemsSource = null;
            lstInFileMacros.ItemsSource = InFileMacros;

            // Load Settings
            if (File.Exists(settingsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(settingsFilePath);
                    var serializer = new JavaScriptSerializer();
                    Settings = serializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
                catch { Settings = new AppSettings(); }
            }
            else Settings = new AppSettings();

            // Initialize UI with Settings
            if (Settings.IsAdvancedMode) rbAdvanced.IsChecked = true;
            else rbSimple.IsChecked = true;

            // Simple
            string simpleProvider = Settings.Provider;
            if (string.IsNullOrEmpty(simpleProvider))
            {
                // Heuristic for legacy settings
                if (Settings.ApiUrl?.Contains("localhost:11434") == true) simpleProvider = "Ollama";
                else if (Settings.ApiUrl?.Contains("localhost:1234") == true) simpleProvider = "LM Studio";
                else simpleProvider = "Custom (OpenAI Compatible)";
            }
            // Handle legacy "OpenAI" or "Custom" saved values
            if (simpleProvider == "OpenAI" || simpleProvider == "Custom") simpleProvider = "Custom (OpenAI Compatible)";
            
            SetComboValue(cmbSimpleProvider, simpleProvider);
            
            txtSimpleApiUrl.Text = Settings.ApiUrl;
            txtSimpleApiKey.Text = Settings.ApiKey;
            cmbSimpleModel.Text = Settings.Model;

            // Advanced
            // Header
            SetComboValue(cmbHeaderProvider, Settings.HeaderDetectionLLM.Provider);
            txtHeaderApiUrl.Text = Settings.HeaderDetectionLLM.ApiUrl;
            txtHeaderApiKey.Text = Settings.HeaderDetectionLLM.ApiKey;
            cmbHeaderModel.Text = Settings.HeaderDetectionLLM.Model;

            // Write
            SetComboValue(cmbWriteProvider, Settings.DataWriteLLM.Provider);
            txtWriteApiUrl.Text = Settings.DataWriteLLM.ApiUrl;
            txtWriteApiKey.Text = Settings.DataWriteLLM.ApiKey;
            cmbWriteModel.Text = Settings.DataWriteLLM.Model;

            // Op
            SetComboValue(cmbOpProvider, Settings.DataOpLLM.Provider);
            txtOpApiUrl.Text = Settings.DataOpLLM.ApiUrl;
            txtOpApiKey.Text = Settings.DataOpLLM.ApiKey;
            cmbOpModel.Text = Settings.DataOpLLM.Model;

            // Vba
            SetComboValue(cmbVbaProvider, Settings.VBASelfHealingLLM.Provider);
            txtVbaApiUrl.Text = Settings.VBASelfHealingLLM.ApiUrl;
            txtVbaApiKey.Text = Settings.VBASelfHealingLLM.ApiKey;
            cmbVbaModel.Text = Settings.VBASelfHealingLLM.Model;
        }

        private void LoadInFileMacros()
        {
            InFileMacros = new List<MacroPreset>();
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWorkbook != null)
                {
                    try
                    {
                        // Check access
                        var proj = app.ActiveWorkbook.VBProject;
                        foreach (dynamic vbComp in proj.VBComponents)
                        {
                            if (vbComp.CodeModule != null)
                            {
                                int count = vbComp.CodeModule.CountOfLines;
                                if (count > 0)
                                {
                                    string code = vbComp.CodeModule.Lines(1, count);
                                    var matches = System.Text.RegularExpressions.Regex.Matches(code, @"Sub\s+(\w+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                    foreach (System.Text.RegularExpressions.Match match in matches)
                                    {
                                        InFileMacros.Add(new MacroPreset
                                        {
                                            Title = match.Groups[1].Value,
                                            Code = code
                                        });
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        InFileMacros.Add(new MacroPreset { Title = "Enable 'Trust access to VBA' to see macros", Code = "" });
                    }
                }
            }
            catch { }
        }

        private void SetComboValue(ComboBox cmb, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            foreach (ComboBoxItem item in cmb.Items)
            {
                if (item.Content.ToString() == value)
                {
                    cmb.SelectedItem = item;
                    break;
                }
            }
        }

        // --- Prompts ---
        private void LstPrompts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstPrompts.SelectedItem is PromptPreset p)
            {
                txtPromptTitle.Text = p.Title;
                txtPromptContent.Text = p.Content;
            }
        }

        private void BtnAddPrompt_Click(object sender, RoutedEventArgs e)
        {
            var newPrompt = new PromptPreset { Title = "New Prompt", Content = "" };
            Prompts.Add(newPrompt);
            RefreshPrompts();
            lstPrompts.SelectedItem = newPrompt;
        }

        private void BtnDeletePrompt_Click(object sender, RoutedEventArgs e)
        {
            if (lstPrompts.SelectedItem is PromptPreset p)
            {
                Prompts.Remove(p);
                RefreshPrompts();
            }
        }

        private void BtnSavePrompt_Click(object sender, RoutedEventArgs e)
        {
            if (lstPrompts.SelectedItem is PromptPreset p)
            {
                p.Title = txtPromptTitle.Text;
                p.Content = txtPromptContent.Text;
                RefreshPrompts();
                SavePromptsToFile();
            }
        }

        private void RefreshPrompts()
        {
            lstPrompts.ItemsSource = null;
            lstPrompts.ItemsSource = Prompts;
        }

        private void SavePromptsToFile()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(Prompts);
            File.WriteAllText(promptsFilePath, json);
        }

        // --- Macros ---
        private void LstMacros_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstMacros.SelectedItem is MacroPreset m)
            {
                txtMacroTitle.Text = m.Title;
                txtMacroCode.Text = m.Code;
                btnSaveMacro.IsEnabled = true;
                btnDeleteMacro.IsEnabled = true;
                
                // Deselect other list to avoid confusion
                if (lstInFileMacros.SelectedIndex != -1)
                {
                    lstInFileMacros.SelectionChanged -= LstInFileMacros_SelectionChanged;
                    lstInFileMacros.SelectedIndex = -1;
                    lstInFileMacros.SelectionChanged += LstInFileMacros_SelectionChanged;
                }
            }
        }

        private void LstInFileMacros_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstInFileMacros.SelectedItem is MacroPreset m)
            {
                txtMacroTitle.Text = m.Title;
                txtMacroCode.Text = m.Code;
                btnSaveMacro.IsEnabled = false; // Cannot edit in-file macros directly here
                btnDeleteMacro.IsEnabled = false;

                // Deselect other list
                if (lstMacros.SelectedIndex != -1)
                {
                    lstMacros.SelectionChanged -= LstMacros_SelectionChanged;
                    lstMacros.SelectedIndex = -1;
                    lstMacros.SelectionChanged += LstMacros_SelectionChanged;
                }
            }
        }

        private void BtnAddMacro_Click(object sender, RoutedEventArgs e)
        {
            var newMacro = new MacroPreset { Title = "New Macro", Code = "Sub NewMacro()\n\nEnd Sub" };
            Macros.Add(newMacro);
            RefreshMacros();
            lstMacros.SelectedItem = newMacro;
        }

        private void BtnDeleteMacro_Click(object sender, RoutedEventArgs e)
        {
            if (lstMacros.SelectedItem is MacroPreset m)
            {
                Macros.Remove(m);
                RefreshMacros();
            }
        }

        private void BtnSaveMacro_Click(object sender, RoutedEventArgs e)
        {
            if (lstMacros.SelectedItem is MacroPreset m)
            {
                m.Title = txtMacroTitle.Text;
                m.Code = txtMacroCode.Text;
                RefreshMacros();
                SaveMacrosToFile();
            }
        }

        private void RefreshMacros()
        {
            lstMacros.ItemsSource = null;
            lstMacros.ItemsSource = Macros;
        }

        private void SaveMacrosToFile()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(Macros);
            File.WriteAllText(macrosFilePath, json);
        }

        // --- LLM ---
        private async void BtnDetectModels_Click(object sender, RoutedEventArgs e)
        {
            lblDetectionStatus.Text = "Detecting...";
            var found = new List<string>();

            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(5);
                
                // Check Ollama
                try
                {
                    var response = await client.GetAsync("http://localhost:11434");
                    if (response.IsSuccessStatusCode) found.Add("Ollama");
                }
                catch { }

                // Check LM Studio
                try
                {
                    var response = await client.GetAsync("http://localhost:1234/v1/models");
                    if (response.IsSuccessStatusCode) found.Add("LM Studio");
                }
                catch { }
            }

            if (found.Count > 0)
            {
                lblDetectionStatus.Text = "Found: " + string.Join(", ", found);
                MessageBox.Show($"Detected local services: {string.Join(", ", found)}.\nYou can now select them in the Provider dropdowns.", "Detection Complete");
            }
            else
            {
                lblDetectionStatus.Text = "No local services found.";
            }
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (grpSimple == null || pnlAdvanced == null) return;

            if (rbSimple.IsChecked == true)
            {
                grpSimple.Visibility = Visibility.Visible;
                pnlAdvanced.Visibility = Visibility.Collapsed;
            }
            else
            {
                grpSimple.Visibility = Visibility.Collapsed;
                pnlAdvanced.Visibility = Visibility.Visible;
            }
        }

        private async void CmbSimpleProvider_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (txtSimpleApiUrl == null) return;
            var item = cmbSimpleProvider.SelectedItem as ComboBoxItem;
            if (item == null) return;

            string provider = item.Content.ToString();
            bool isLocal = provider == "Ollama" || provider == "LM Studio";
            
            if (provider == "Ollama") txtSimpleApiUrl.Text = "http://localhost:11434/v1";
            else if (provider == "LM Studio") txtSimpleApiUrl.Text = "http://localhost:1234/v1";

            if (pnlSimpleApiKey != null) pnlSimpleApiKey.Visibility = isLocal ? Visibility.Collapsed : Visibility.Visible;

            await PopulateModelList(provider, txtSimpleApiUrl.Text, cmbSimpleModel);
        }

        private async void CmbAdvancedProvider_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            if (cmb == null) return;
            
            var item = cmb.SelectedItem as ComboBoxItem;
            if (item == null) return;

            string tag = cmb.Tag.ToString();
            string provider = item.Content.ToString();
            bool isLocal = provider == "Ollama" || provider == "LM Studio";
            string url = "";

            if (provider == "Ollama") url = "http://localhost:11434/v1";
            else if (provider == "LM Studio") url = "http://localhost:1234/v1";

            ComboBox targetModelCmb = null;
            TextBox targetUrlBox = null;
            StackPanel targetApiKeyPanel = null;

            if (tag == "Header") 
            {
                targetUrlBox = txtHeaderApiUrl;
                targetModelCmb = cmbHeaderModel;
                targetApiKeyPanel = pnlHeaderApiKey;
            }
            else if (tag == "Write") 
            {
                targetUrlBox = txtWriteApiUrl;
                targetModelCmb = cmbWriteModel;
                targetApiKeyPanel = pnlWriteApiKey;
            }
            else if (tag == "Op") 
            {
                targetUrlBox = txtOpApiUrl;
                targetModelCmb = cmbOpModel;
                targetApiKeyPanel = pnlOpApiKey;
            }
            else if (tag == "Vba") 
            {
                targetUrlBox = txtVbaApiUrl;
                targetModelCmb = cmbVbaModel;
                targetApiKeyPanel = pnlVbaApiKey;
            }

            if (targetUrlBox != null && !string.IsNullOrEmpty(url))
            {
                targetUrlBox.Text = url;
            }
            
            if (targetApiKeyPanel != null)
            {
                targetApiKeyPanel.Visibility = isLocal ? Visibility.Collapsed : Visibility.Visible;
            }

            if (targetModelCmb != null && targetUrlBox != null)
            {
                await PopulateModelList(provider, targetUrlBox.Text, targetModelCmb);
            }
        }

        private async Task PopulateModelList(string provider, string apiUrl, ComboBox targetCmb)
        {
            if (provider != "Ollama" && provider != "LM Studio") return;
            
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(5); // Increased timeout
                    string requestUrl = apiUrl.TrimEnd('/') + "/models";
                    
                    var response = await client.GetAsync(requestUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        var serializer = new JavaScriptSerializer();
                        dynamic result = serializer.Deserialize<dynamic>(json);
                        
                        var models = new List<string>();
                        
                        // OpenAI format (LM Studio & Ollama v1)
                        if (result is Dictionary<string, object> dict && dict.ContainsKey("data"))
                        {
                            var data = dict["data"] as System.Collections.IEnumerable;
                            if (data != null)
                            {
                                foreach (dynamic model in data)
                                {
                                    if (model is Dictionary<string, object> mDict && mDict.ContainsKey("id"))
                                        models.Add(mDict["id"].ToString());
                                }
                            }
                        }
                        
                        if (models.Count > 0)
                        {
                            targetCmb.ItemsSource = models;
                            if (targetCmb.Items.Count > 0 && string.IsNullOrEmpty(targetCmb.Text))
                            {
                                targetCmb.SelectedIndex = 0;
                            }
                        }
                    }
                }
            }
            catch 
            { 
                // Silent fail
            }
        }

        // --- Main Buttons ---
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            // Save Settings
            Settings.IsAdvancedMode = rbAdvanced.IsChecked == true;

            // Simple
            Settings.Provider = (cmbSimpleProvider.SelectedItem as ComboBoxItem)?.Content.ToString();
            Settings.ApiUrl = txtSimpleApiUrl.Text;
            Settings.ApiKey = txtSimpleApiKey.Text;
            Settings.Model = cmbSimpleModel.Text;

            // Advanced
            Settings.HeaderDetectionLLM.Provider = (cmbHeaderProvider.SelectedItem as ComboBoxItem)?.Content.ToString();
            Settings.HeaderDetectionLLM.ApiUrl = txtHeaderApiUrl.Text;
            Settings.HeaderDetectionLLM.ApiKey = txtHeaderApiKey.Text;
            Settings.HeaderDetectionLLM.Model = cmbHeaderModel.Text;

            Settings.DataWriteLLM.Provider = (cmbWriteProvider.SelectedItem as ComboBoxItem)?.Content.ToString();
            Settings.DataWriteLLM.ApiUrl = txtWriteApiUrl.Text;
            Settings.DataWriteLLM.ApiKey = txtWriteApiKey.Text;
            Settings.DataWriteLLM.Model = cmbWriteModel.Text;

            Settings.DataOpLLM.Provider = (cmbOpProvider.SelectedItem as ComboBoxItem)?.Content.ToString();
            Settings.DataOpLLM.ApiUrl = txtOpApiUrl.Text;
            Settings.DataOpLLM.ApiKey = txtOpApiKey.Text;
            Settings.DataOpLLM.Model = cmbOpModel.Text;

            Settings.VBASelfHealingLLM.Provider = (cmbVbaProvider.SelectedItem as ComboBoxItem)?.Content.ToString();
            Settings.VBASelfHealingLLM.ApiUrl = txtVbaApiUrl.Text;
            Settings.VBASelfHealingLLM.ApiKey = txtVbaApiKey.Text;
            Settings.VBASelfHealingLLM.Model = cmbVbaModel.Text;

            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(Settings);
            File.WriteAllText(settingsFilePath, json);

            SettingsSaved?.Invoke();
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
