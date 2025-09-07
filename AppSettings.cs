using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MirrorAudio.AppContextApp
{
    public sealed class AppSettings
    {
        public string? InputDeviceId { get; set; } // null = 默认系统输出(环回)
        public string? MainDeviceId { get; set; }   // 主通道（期望SPDIF）
        public string? AuxDeviceId { get; set; }    // 副通道

        public bool MainExclusive { get; set; } = true;
        public bool MainRaw { get; set; } = false;           // RAW直通（绕开APM）
        public bool MainForceFormat { get; set; } = true;    // 强制使用下方格式
        public int MainSampleRate { get; set; } = 192000;
        public int MainBits { get; set; } = 24;
        public int MainBufferMs { get; set; } = 12;

        public bool AuxExclusive { get; set; } = true;
        public bool AuxRaw { get; set; } = false;
        public bool AuxForceFormat { get; set; } = true;
        public int AuxSampleRate { get; set; } = 44100;
        public int AuxBits { get; set; } = 16;
        public int AuxBufferMs { get; set; } = 120;

        public bool AutoStart { get; set; } = false;
        public bool EnableLogging { get; set; } = true;

        [JsonIgnore]
        public static string Dir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MirrorAudio");
        [JsonIgnore] public static string PathFile => System.IO.Path.Combine(Dir, "config.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(PathFile))
                {
                    var json = File.ReadAllText(PathFile);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch { }
            return new AppSettings();
        }

        public void Save()
        {
            Directory.CreateDirectory(Dir);
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(PathFile, json);
        }
    }
}