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
        
        // New Settings
        public bool IsAdvancedMode { get; set; }
        public LLMConfig HeaderDetectionLLM { get; set; }
        public LLMConfig DataWriteLLM { get; set; }
        public LLMConfig DataOpLLM { get; set; }
        public LLMConfig VBASelfHealingLLM { get; set; }

        public AppSettings()
        {
            HeaderDetectionLLM = new LLMConfig();
            DataWriteLLM = new LLMConfig();
            DataOpLLM = new LLMConfig();
            VBASelfHealingLLM = new LLMConfig();
        }
    }

    public class LLMConfig
    {
        public string Provider { get; set; } // "OpenAI", "Ollama", "LMStudio"
        public string ApiUrl { get; set; }
        public string ApiKey { get; set; }
        public string Model { get; set; }
    }

    public class HeaderInfo
    {
        public string HeaderContent { get; set; }
        public int HeaderRows { get; set; }
    }
}
