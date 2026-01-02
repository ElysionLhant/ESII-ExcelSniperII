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
}
