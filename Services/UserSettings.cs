using System;
using System.IO;
using System.Text.Json;

namespace SmartDocProcessor.WPF.Services
{
    public class UserSettings
    {
        public string DefaultFontFamily { get; set; } = "Malgun Gothic";
        public int DefaultFontSize { get; set; } = 14;
        public string DefaultColor { get; set; } = "#FF000000"; // Black
        public bool DefaultIsBold { get; set; } = false;

        private static string _filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "user_settings.json");

        public static UserSettings Load()
        {
            if (File.Exists(_filePath))
            {
                try
                {
                    string json = File.ReadAllText(_filePath);
                    return JsonSerializer.Deserialize<UserSettings>(json) ?? new UserSettings();
                }
                catch { }
            }
            return new UserSettings();
        }

        public void Save()
        {
            try
            {
                string json = JsonSerializer.Serialize(this);
                File.WriteAllText(_filePath, json);
            }
            catch { }
        }
    }
}