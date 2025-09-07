using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;
using NAudio.Wave;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Load config
            _cfg = LoadConfig();

            // UI
            _form = new SettingsForm(
                getStatus: GetStatusSnapshot,
                saveConfig: SaveConfig,
                getConfig: () => _cfg,
                applyAndRestart: StartOrRestart,
                cleanup: Cleanup);

            Application.ApplicationExit += (_, __) => Cleanup();
            _form.Shown += (_, __) => { StartOrRestart(); _form.RenderStatus(); };
            Application.Run(_form);
        }

        // === Config ===
        public static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mirror_audio.settings.json");

        [DataContract]
        public class AppSettings
        {
            [DataMember] public int MainRate = 192000;
            [DataMember] public int MainBits = 24;
            [DataMember] public int MainBufMs = 9;
            [DataMember] public int AuxRate = 44100;
            [DataMember] public int AuxBits = 16;
            [DataMember] public int AuxBufMs = 120;
            [DataMember] public bool MainExclusive = true;
            [DataMember] public bool AuxExclusive = false;
            [DataMember] public bool MainEventDriven = true;
            [DataMember] public bool AuxEventDriven = false;

            // New fields
            [DataMember] public bool MainStrictFormat = false;
            [DataMember] public bool AuxStrictFormat = false;
            [DataMember] public int MainAlignMultiple = 0; // 0 = 自动（按最小周期向上取整到整数倍）；>0 = 强制 N × 最小周期
            [DataMember] public int AuxAlignMultiple = 0;
        }

        static AppSettings _cfg = new AppSettings();

        static AppSettings LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    using (var fs = File.OpenRead(ConfigPath))
                    {
                        var ser = new DataContractJsonSerializer(typeof(AppSettings));
                        return (AppSettings)ser.ReadObject(fs);
                    }
                }
            }
            catch { /* ignore */ }
            return new AppSettings();
        }

        static void SaveConfig(AppSettings cfg)
        {
            try
            {
                using (var fs = File.Create(ConfigPath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    ser.WriteObject(fs, cfg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存配置失败: " + ex.Message);
            }
        }

        // === Audio state ===

        static MMDevice _outMain;
        static MMDevice _outAux;

        static WasapiOut _mainOut;
        static WasapiOut _auxOut;

        static BufferedWaveProvider _bufMain;
        static BufferedWaveProvider _bufAux;

        static IWaveProvider _srcMain;
        static IWaveProvider _srcAux;

        static MediaFoundationResampler _resMain;
        static MediaFoundationResampler _resAux;

        static SettingsForm _form;

        // Status
        static bool _mainIsExclusive;
        static bool _auxIsExclusive;

        static double _defMainMs = 3.0; // default period for event mode rough baseline
        static double _defAuxMs = 3.0;
        static double _minMainMs = 3.0; // reported minimum period (ms)
        static double _minAuxMs = 3.0;

        static int _mainBufEffectiveMs;
        static int _auxBufEffectiveMs;

        static string _mainFmtStr = "";
        static string _auxFmtStr = "";

        static bool _mainResampling;
        static bool _auxResampling;
        static bool _mainNoSRC;
        static bool _auxNoSRC;

        static WaveFormat _mainTargetFmt;
        static WaveFormat _auxTargetFmt;

        // === Public helpers used by UI ===
        public class StatusSnapshot
        {
            public string MainDevice;
            public string AuxDevice;
            public string MainMode; // Exclusive/Shared + Event/Timer
            public string AuxMode;
            public string MainFmt;  // e.g., 192k/24/2ch (容器)
            public string AuxFmt;
            public string MainBitDepthMap; // 容器位深 → 线缆位深
            public string AuxBitDepthMap;
            public string MainBuf;
            public string AuxBuf;
            public string MainResample; // 是/否
            public string AuxResample;
        }

        static StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot();
            s.MainDevice = _outMain?.FriendlyName ?? "(未选择)";
            s.AuxDevice = _outAux?.FriendlyName ?? "(未选择)";
            s.MainMode = (_mainIsExclusive ? "独占" : "共享") + " / " + (_cfg.MainEventDriven ? "事件驱动" : "轮询");
            s.AuxMode = (_auxIsExclusive ? "独占" : "共享") + " / " + (_cfg.AuxEventDriven ? "事件驱动" : "轮询");
            s.MainFmt = _mainFmtStr;
            s.AuxFmt = _auxFmtStr;
            s.MainBitDepthMap = BuildBitDepthMap(_outMain, _mainTargetFmt, _cfg.MainBits);
            s.AuxBitDepthMap = BuildBitDepthMap(_outAux, _auxTargetFmt, _cfg.AuxBits);
            s.MainBuf = _mainBufEffectiveMs > 0 ? $"{_mainBufEffectiveMs} ms" : "-";
            s.AuxBuf = _auxBufEffectiveMs > 0 ? $"{_auxBufEffectiveMs} ms" : "-";
            s.MainResample = _mainResampling ? "是" : (_mainNoSRC ? "直通" : "否");
            s.AuxResample = _auxResampling ? "是" : (_auxNoSRC ? "直通" : "否");
            return s;
        }

        static string Fmt(WaveFormat w)
        {
            if (w == null) return "-";
            var bits = w.Encoding == WaveFormatEncoding.IeeeFloat ? "32f" : w.BitsPerSample.ToString();
            return $"{w.SampleRate / 1000.0:0.#}k/{bits}/{w.Channels}ch";
        }

        static string BuildBitDepthMap(MMDevice dev, WaveFormat containerFmt, int requestedBits)
        {
            if (dev == null || containerFmt == null) return "-";
            string container = containerFmt.Encoding == WaveFormatEncoding.IeeeFloat ? "32f" : containerFmt.BitsPerSample.ToString();
            // 线缆有效位深估算：SPDIF 等常见为 24bit，有时容器 32 仅为填充
            string name = dev.FriendlyName?.ToLowerInvariant() ?? "";
            bool isSpdif = name.Contains("spdif") || name.Contains("s/pdif") || name.Contains("digital") || name.Contains("optical") || name.Contains("光纤");
            string cable = container;
            string note = "";
            if (isSpdif)
            {
                if ((containerFmt.BitsPerSample == 32 || containerFmt.Encoding == WaveFormatEncoding.IeeeFloat) && requestedBits == 24)
                {
                    cable = "24";
                    note = "（填充）";
                }
            }
            return $"容器位深 {container} → 线缆位深 {cable}{note}";
        }

        // === Core start/stop ===
        static void Cleanup()
        {
            try
            {
                _mainOut?.Stop(); _mainOut?.Dispose(); _mainOut = null;
                _auxOut?.Stop(); _auxOut?.Dispose(); _auxOut = null;
                _resMain?.Dispose(); _resMain = null;
                _resAux?.Dispose(); _resAux = null;
                _outMain?.Dispose(); _outMain = null;
                _outAux?.Dispose(); _outAux = null;
            }
            catch { }
        }

        static MMDevice GetDefaultRender()
        {
            var en = new MMDeviceEnumerator();
            return en.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        }

        static WaveFormat MakeWanted(int rate, int bits, int ch = 2)
        {
            if (bits == 32) return WaveFormat.CreateIeeeFloatWaveFormat(rate, ch); // 在很多链路上更好协商
            return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, ch, rate * ch * (bits / 8), ch * (bits / 8), bits);
        }

        static bool SupportsExclusive(MMDevice dev, WaveFormat fmt)
        {
            try
            {
                using (var ac = dev.AudioClient)
                {
                    ac.IsFormatSupported(AudioClientShareMode.Exclusive, fmt, out var _);
                    return true;
                }
            }
            catch { return false; }
        }

        // alignMultiple: 0=自动，>0 时强制 N×最小周期；exclusive 标志决定底线
        static int Buf(int want, bool exclusive, double defMs, double minMs = 0, int alignMultiple = 0)
        {
            int ms = want;
            if (exclusive)
            {
                // 稳定性底线：至少 >= 3 × 默认周期
                int floor = (int)Math.Ceiling(defMs * 3.0);
                if (ms < floor) ms = floor;

                if (minMs > 0)
                {
                    if (alignMultiple > 0)
                    {
                        ms = (int)Math.Ceiling(alignMultiple * minMs);
                    }
                    else
                    {
                        double k = Math.Ceiling(ms / minMs);
                        ms = (int)Math.Ceiling(k * minMs);
                    }
                }
            }
            else
            {
                int floor = (int)Math.Ceiling(defMs * 2.0);
                if (ms < floor) ms = floor;
            }
            return ms;
        }

        static WasapiOut CreateOut(MMDevice dev, bool exclusive, bool eventDriven, int bufferMs, IWaveProvider src)
        {
            var share = exclusive ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared;
            var outMode = eventDriven ? AudioClientStreamFlags.EventCallback : AudioClientStreamFlags.None;
            var wo = new WasapiOut(dev, share, eventDriven ? true : false, bufferMs);
            try
            {
                wo.Init(src);
                return wo;
            }
            catch
            {
                wo.Dispose();
                return null;
            }
        }

        static void StartOrRestart()
        {
            Cleanup();

            // pick devices
            _outMain = GetDefaultRender();
            _outAux = _outMain; // 若有副设备选择逻辑，可在此替换

            // buffers
            _bufMain = new BufferedWaveProvider(WaveFormat.CreateIeeeFloatWaveFormat(_cfg.MainRate, 2)) { DiscardOnBufferOverflow = true };
            _bufAux = new BufferedWaveProvider(WaveFormat.CreateIeeeFloatWaveFormat(_cfg.AuxRate, 2)) { DiscardOnBufferOverflow = true };

            // MAIN
            _mainIsExclusive = _cfg.MainExclusive;
            _mainTargetFmt = MakeWanted(_cfg.MainRate, _cfg.MainBits, 2);
            _srcMain = _bufMain;
            _resMain = null;
            _mainResampling = false; _mainNoSRC = true;

            WaveFormat inMainFmt = _bufMain.WaveFormat;
            bool needResampleMain = (inMainFmt.SampleRate != _mainTargetFmt.SampleRate) || (inMainFmt.Channels != _mainTargetFmt.Channels) ||
                                    ((inMainFmt.Encoding == WaveFormatEncoding.IeeeFloat ? 32 : inMainFmt.BitsPerSample) != (_mainTargetFmt.Encoding == WaveFormatEncoding.IeeeFloat ? 32 : _mainTargetFmt.BitsPerSample));
            if (needResampleMain)
            {
                _resMain = new MediaFoundationResampler(_bufMain, _mainTargetFmt) { ResamplerQuality = 50 };
                _srcMain = _resMain;
                _mainResampling = true; _mainNoSRC = false;
            }

            int msMain = Buf(_cfg.MainBufMs, _cfg.MainExclusive, _defMainMs, _minMainMs, _cfg.MainAlignMultiple);
            _mainOut = CreateOut(_outMain, _cfg.MainExclusive, _cfg.MainEventDriven, msMain, _srcMain);

            // 处理 24->32 回退：仅当未开启严格格式时允许
            if (_mainOut == null && !_cfg.MainStrictFormat && _cfg.MainExclusive && _cfg.MainBits == 24)
            {
                var fmt32 = WaveFormat.CreateIeeeFloatWaveFormat(_cfg.MainRate, 2);
                bool try32 = SupportsExclusive(_outMain, fmt32);
                if (try32)
                {
                    _mainTargetFmt = fmt32;
                    if (_resMain != null) { _resMain.Dispose(); _resMain = null; }
                    bool needRateChange = (inMainFmt.SampleRate != fmt32.SampleRate) || (inMainFmt.Channels != fmt32.Channels);
                    _srcMain = _bufMain;
                    if (needRateChange)
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, fmt32) { ResamplerQuality = 50 };
                        _srcMain = _resMain;
                        _mainResampling = true; _mainNoSRC = false;
                    }
                    int ms32 = Buf(_cfg.MainBufMs, _cfg.MainExclusive, _defMainMs, _minMainMs, _cfg.MainAlignMultiple);
                    _mainOut = CreateOut(_outMain, _cfg.MainExclusive, _cfg.MainEventDriven, ms32, _srcMain);
                    if (_mainOut != null) msMain = ms32;
                }
            }

            if (_cfg.MainExclusive && _cfg.MainStrictFormat && _mainOut == null)
            {
                MessageBox.Show("主通道独占失败（严格格式）。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _mainBufEffectiveMs = _mainOut != null ? msMain : 0;
            _mainFmtStr = Fmt(_mainTargetFmt);

            if (_mainOut != null) _mainOut.Play();

            // AUX
            _auxIsExclusive = _cfg.AuxExclusive;
            _auxTargetFmt = MakeWanted(_cfg.AuxRate, _cfg.AuxBits, 2);
            _srcAux = _bufAux;
            _resAux = null;
            _auxResampling = false; _auxNoSRC = true;

            WaveFormat inAuxFmt = _bufAux.WaveFormat;
            bool needResampleAux = (inAuxFmt.SampleRate != _auxTargetFmt.SampleRate) || (inAuxFmt.Channels != _auxTargetFmt.Channels) ||
                                   ((inAuxFmt.Encoding == WaveFormatEncoding.IeeeFloat ? 32 : inAuxFmt.BitsPerSample) != (_auxTargetFmt.Encoding == WaveFormatEncoding.IeeeFloat ? 32 : _auxTargetFmt.BitsPerSample));
            if (needResampleAux)
            {
                _resAux = new MediaFoundationResampler(_bufAux, _auxTargetFmt) { ResamplerQuality = 50 };
                _srcAux = _resAux;
                _auxResampling = true; _auxNoSRC = false;
            }

            int msAux = Buf(_cfg.AuxBufMs, _cfg.AuxExclusive, _defAuxMs, _minAuxMs, _cfg.AuxAlignMultiple);
            _auxOut = CreateOut(_outAux, _cfg.AuxExclusive, _cfg.AuxEventDriven, msAux, _srcAux);

            if (_auxOut == null && !_cfg.AuxStrictFormat && _cfg.AuxExclusive && _cfg.AuxBits == 24)
            {
                var fmt32 = WaveFormat.CreateIeeeFloatWaveFormat(_cfg.AuxRate, 2);
                bool try32 = SupportsExclusive(_outAux, fmt32);
                if (try32)
                {
                    _auxTargetFmt = fmt32;
                    if (_resAux != null) { _resAux.Dispose(); _resAux = null; }
                    bool needRateChange = (inAuxFmt.SampleRate != fmt32.SampleRate) || (inAuxFmt.Channels != fmt32.Channels);
                    _srcAux = _bufAux;
                    if (needRateChange)
                    {
                        _resAux = new MediaFoundationResampler(_bufAux, fmt32) { ResamplerQuality = 50 };
                        _srcAux = _resAux;
                        _auxResampling = true; _auxNoSRC = false;
                    }
                    int ms32 = Buf(_cfg.AuxBufMs, _cfg.AuxExclusive, _defAuxMs, _minAuxMs, _cfg.AuxAlignMultiple);
                    _auxOut = CreateOut(_outAux, _cfg.AuxExclusive, _cfg.AuxEventDriven, ms32, _srcAux);
                    if (_auxOut != null) msAux = ms32;
                }
            }

            if (_cfg.AuxExclusive && _cfg.AuxStrictFormat && _auxOut == null)
            {
                MessageBox.Show("副通道独占失败（严格格式）。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _auxBufEffectiveMs = _auxOut != null ? msAux : 0;
            _auxFmtStr = Fmt(_auxTargetFmt);

            if (_auxOut != null) _auxOut.Play();

            _form?.RenderStatus();
        }
    }
}
