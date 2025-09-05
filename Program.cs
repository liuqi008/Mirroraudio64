using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.MediaFoundation;
using NAudio.Wave;

// 解决 Timer 二义性
using WinFormsTimer = System.Windows.Forms.Timer;

namespace MirrorAudio
{
    // 轻量日志（默认关）
    static class Logger
    {
        public static bool Enabled;
        static readonly string LogDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        static string PathOf(string name) { try { if (!Directory.Exists(LogDir)) Directory.CreateDirectory(LogDir); } catch { } return System.IO.Path.Combine(LogDir, name); }
        public static void Info(string msg)
        {
            if (!Enabled) return;
            try { File.AppendAllText(PathOf("MirrorAudio.log"), $"[{DateTime.Now:HH:mm:ss}] {msg}\r\n"); } catch { }
        }
        public static void Crash(string where, Exception ex)
        {
            if (ex == null) return;
            try { File.AppendAllText(PathOf("MirrorAudio.crash.log"), $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n"); } catch { }
        }
    }

    [DataContract]
    public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }

    [DataContract]
    public enum SyncModeOption { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    [DataContract]
    public sealed class AppSettings
    {
        // 设备选择
        [DataMember] public string InputDeviceId;
        [DataMember] public string MainDeviceId;
        [DataMember] public string AuxDeviceId;

        // 主通道（追求低延迟）
        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto;
        [DataMember] public SyncModeOption  MainSync  = SyncModeOption.Auto;
        [DataMember] public int MainRate  = 192000; // 独占才用
        [DataMember] public int MainBits  = 24;     // 独占才用（24 不通再试 32 容器）
        [DataMember] public int MainBufMs = 12;

        // 副通道（推流/省电）
        [DataMember] public ShareModeOption AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption  AuxSync  = SyncModeOption.Event;
        [DataMember] public int AuxRate  = 48000; // 独占才用
        [DataMember] public int AuxBits  = 16;    // 独占才用
        [DataMember] public int AuxBufMs = 180;

        // 其他
        [DataMember] public bool AutoStart = false;
        [DataMember] public bool EnableLogging = false;
    }

    public sealed class StatusSnapshot
    {
        public bool   Running;
        public string InputRole;
        public string InputFormat;
        public string InputDevice;
        public string MainDevice, AuxDevice;
        public string MainMode,   AuxMode;
        public string MainSync,   AuxSync;
        public string MainFormat, AuxFormat;
        public int    MainBufferMs, AuxBufferMs;
    }

    static class Config
    {
        static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
        static readonly string FilePath = Path.Combine(Dir, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new AppSettings();
                using (var fs = File.OpenRead(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    return (AppSettings)ser.ReadObject(fs);
                }
            }
            catch { return new AppSettings(); }
        }

        public static void Save(AppSettings s)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                using (var fs = File.Create(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    ser.WriteObject(fs, s);
                }
            }
            catch { }
        }
    }

    static class Program
    {
        static Mutex _single;

        [STAThread]
        static void Main()
        {
            bool created;
            _single = new Mutex(true, "Global\\MirrorAudio_{5C5C126B-B5D3-4F93-A0AB-5E4E5D94FB5B}", out created);
            if (!created) return;

            Application.ThreadException += (s, e) => Logger.Crash("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Crash("NonUI", e.ExceptionObject as Exception);

            try { MediaFoundationApi.Startup(); } catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) Application.Run();

            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip { ShowCheckMargin = false, ShowImageMargin = true };
        readonly WinFormsTimer _debounce = new WinFormsTimer { Interval = 420 };

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        // 设备
        MMDevice _inDev, _outMain, _outAux;

        // 输入与两路输出
        IWaveIn _capture;
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux;
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;

        // 状态
        bool _running;
        bool _mainIsExclusive, _auxIsExclusive;
        bool _mainEventUsed, _auxEventUsed;
        int  _mainBufMsEff, _auxBufMsEff;
        string _inRole = "-", _inFmt = "-", _inName = "-", _mainFmt = "-", _auxFmt = "-";

        public TrayApp()
        {
            Logger.Enabled = _cfg.EnableLogging;
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            // 托盘
            try
            {
                if (File.Exists("MirrorAudio.ico")) _tray.Icon = new Icon("MirrorAudio.ico");
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true;
            _tray.Text = "MirrorAudio";

            _menu.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripMenuItem("启动/重启(&S)", To16(SystemIcons.Information), (s,e)=> StartOrRestart()),
                new ToolStripMenuItem("停止(&T)",     To16(SystemIcons.Hand),        (s,e)=> Stop()),
                new ToolStripMenuItem("设置(&G)...",  To16(SystemIcons.Information), OnSettings),
                new ToolStripMenuItem("日志目录",      To16(SystemIcons.Asterisk),    (s,e)=> OpenLogDir()),
                new ToolStripMenuItem("退出(&X)",     To16(SystemIcons.Error),       (s,e)=> { Stop(); Application.Exit(); })
            });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s,e)=> { _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        static Bitmap To16(Icon ico)
        {
            try
            {
                using (var bmp = ico.ToBitmap())
                {
                    var img = new Bitmap(16, 16);
                    using (var g = Graphics.FromImage(img)) g.DrawImage(bmp, new Rectangle(0, 0, 16, 16));
                    return img;
                }
            } catch { return null; }
        }

        void OpenLogDir()
        {
            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            try { if (!Directory.Exists(dir)) Directory.CreateDirectory(dir); } catch { }
            try { Process.Start("explorer.exe", dir); } catch { }
        }

        void OnSettings(object sender, EventArgs e)
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result;
                    Logger.Enabled = _cfg.EnableLogging;
                    Config.Save(_cfg);
                    EnsureAutoStart(_cfg.AutoStart);
                    StartOrRestart();
                }
            }
        }

        // —— 设备事件：事件驱动热插拔自愈 —— //
        public void OnDeviceStateChanged(string id, DeviceState st) { if (IsRel(id)) Debounce(); }
        public void OnDeviceAdded(string id) { if (IsRel(id)) Debounce(); }
        public void OnDeviceRemoved(string id) { if (IsRel(id)) Debounce(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string id)
        {
            if (flow == DataFlow.Render && role == Role.Multimedia && string.IsNullOrEmpty(_cfg?.InputDeviceId)) Debounce();
        }
        public void OnPropertyValueChanged(string id, PropertyKey key) { if (IsRel(id)) Debounce(); }
        bool IsRel(string id)
        {
            if (string.IsNullOrEmpty(id) || _cfg == null) return false;
            return string.Equals(id, _cfg.InputDeviceId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.MainDeviceId,  StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.AuxDeviceId,   StringComparison.OrdinalIgnoreCase);
        }
        void Debounce() { _debounce.Stop(); _debounce.Start(); }

        // —— 主流程 —— //
        void StartOrRestart()
        {
            Stop(); // 先停再启，确保干净
            if (_mm == null) _mm = new MMDeviceEnumerator();

            // 输入优先用配置；未选时用默认渲染环回
            _inDev   = FindById(_cfg.InputDeviceId, DataFlow.Capture) ?? FindById(_cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_cfg.MainDeviceId,  DataFlow.Render);
            _outAux  = FindById(_cfg.AuxDeviceId,   DataFlow.Render);

            if (_outMain == null || _outAux == null)
            {
                MessageBox.Show("请先在“设置”里选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 输入：录音 or 环回
            WaveFormat inFmt;
            if (_inDev.DataFlow == DataFlow.Capture)
            {
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                _capture = cap; inFmt = cap.WaveFormat; _inRole = "录音";
            }
            else
            {
                var cap = new WasapiLoopbackCapture(_inDev);
                _capture = cap; inFmt = cap.WaveFormat; _inRole = "环回";
            }
            _inName = _inDev.FriendlyName;
            _inFmt  = Fmt(inFmt);

            // 两路缓冲（容量适中，减少 GC 压力）
            _bufMain = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow = true, ReadFully = true,
                BufferDuration = TimeSpan.FromMilliseconds(Math.max(_cfg.MainBufMs*6, 120)) };
            _bufAux  = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow = true, ReadFully = true,
                BufferDuration = TimeSpan.FromMilliseconds(Math.max(_cfg.AuxBufMs*4, 150)) };

            // 主通道初始化
            InitOne(out _mainOut, out _srcMain, out _resMain,
                    _outMain, true,
                    _cfg.MainShare, _cfg.MainSync, _cfg.MainBufMs,
                    _cfg.MainRate, _cfg.MainBits,
                    inFmt,
                    out _mainIsExclusive, out _mainEventUsed, out _mainBufMsEff, out _mainFmt);

            if (_mainOut == null) { CleanupCreated(); MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            // 副通道初始化
            InitOne(out _auxOut, out _srcAux, out _resAux,
                    _outAux, false,
                    _cfg.AuxShare, _cfg.AuxSync, _cfg.AuxBufMs,
                    _cfg.AuxRate, _cfg.AuxBits,
                    inFmt,
                    out _auxIsExclusive, out _auxEventUsed, out _auxBufMsEff, out _auxFmt);

            if (_auxOut == null) { CleanupCreated(); MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            // 绑定与启动
            _capture.DataAvailable += OnCapData;
            _capture.RecordingStopped += (s,e)=> { try { _bufMain?.ClearBuffer(); _bufAux?.ClearBuffer(); } catch { } };

            try
            {
                _mainOut.Play();
                _auxOut.Play();
                _capture.StartRecording();
                _running = true;
                if (Logger.Enabled)
                {
                    Logger.Info($"Main: {(_mainIsExclusive?"独占":"共享")} | {(_mainEventUsed?"事件":"轮询")} | buf={_mainBufMsEff}ms | {(_mainFmt ?? "-")}");
                    Logger.Info($"Aux : {(_auxIsExclusive ?"独占":"共享")} | {(_auxEventUsed ?"事件":"轮询")} | buf={_auxBufMsEff}ms | {(_auxFmt  ?? "-")}");
                }
            }
            catch (Exception ex)
            {
                Logger.Crash("Start", ex);
                MessageBox.Show("启动失败：" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void OnCapData(object sender, WaveInEventArgs e)
        {
            var buf = e.Buffer; int len = e.BytesRecorded;
            _bufMain.AddSamples(buf, 0, len);
            _bufAux .AddSamples(buf, 0, len);
        }

        // 初始化一条输出（主/副共用）
        void InitOne(out WasapiOut player, out IWaveProvider src, out MediaFoundationResampler resampler,
                     MMDevice outDev, bool isMain,
                     ShareModeOption shareOpt, SyncModeOption syncOpt, int bufMs,
                     int wantRate, int wantBits,
                     WaveFormat inputFormat,
                     out bool exclusive, out bool eventUsed, out int bufMsEff, out string fmtStr)
        {
            player = null; src = null; resampler = null;
            exclusive = false; eventUsed = false; bufMsEff = bufMs; fmtStr = "-";

            if (outDev == null) return;

            // 冲突：环回同一输出设备时，不允许独占
            bool loopbackConflict = (_inDev.DataFlow == DataFlow.Render) &&
                                    string.Equals(_inDev.ID, outDev.ID, StringComparison.OrdinalIgnoreCase);
            bool wantExclusive = (shareOpt == ShareModeOption.Exclusive || shareOpt == ShareModeOption.Auto) && !loopbackConflict;

            // 先用源缓冲
            src = isMain ? (IWaveProvider)_bufMain : (IWaveProvider)_bufAux;

            // 尝试独占优先（若允许）
            if (wantExclusive)
            {
                var desired = new WaveFormat(wantRate, wantBits, 2);
                if (!FormatsEqual(inputFormat, desired))
                {
                    // 仅在需要时插入重采样（质量适中，主通道略高）
                    resampler = new MediaFoundationResampler(src, desired) { ResamplerQuality = isMain ? 50 : 40 };
                    src = resampler;
                }

                int msEx = SafeBuf(bufMs, true);
                player = CreateOutWithPolicy(outDev, AudioClientShareMode.Exclusive, syncOpt, msEx, src, out eventUsed);
                if (player == null && wantBits == 24)
                {
                    // 24 独占不通时，试 32-bit 容器（24-in-32）
                    var desired32 = new WaveFormat(wantRate, 32, 2);
                    resampler?.Dispose(); resampler = null;
                    if (!FormatsEqual(inputFormat, desired32))
                    {
                        resampler = new MediaFoundationResampler(isMain ? _bufMain : _bufAux, desired32) { ResamplerQuality = isMain ? 50 : 40 };
                        src = resampler;
                    }
                    player = CreateOutWithPolicy(outDev, AudioClientShareMode.Exclusive, syncOpt, msEx, src, out eventUsed);
                    if (player != null) { desired = desired32; }
                }

                if (player != null)
                {
                    exclusive = true; bufMsEff = msEx;
                    fmtStr = Fmt(src.WaveFormat); // 直接使用 IWaveProvider 的 WaveFormat
                }
                else if (shareOpt == ShareModeOption.Exclusive)
                {
                    // 强制独占失败则直接返回失败
                    return;
                }
                else
                {
                    // 独占失败继续走共享
                    resampler?.Dispose(); resampler = null; src = isMain ? (IWaveProvider)_bufMain : (IWaveProvider)_bufAux;
                }
            }

            if (player == null)
            {
                // 共享：由系统混音，采样率/位深以系统为准，无需主动重采样
                int msSh = SafeBuf(bufMs, false);
                player = CreateOutWithPolicy(outDev, AudioClientShareMode.Shared, syncOpt, msSh, src, out eventUsed);
                if (player == null) return;
                exclusive = false; bufMsEff = msSh;
                try { fmtStr = Fmt(outDev.AudioClient.MixFormat); } catch { fmtStr = "系统混音"; }
            }
        }

        // 安全缓冲（轻量规则）
        static int SafeBuf(int ms, bool exclusive)
        {
            if (exclusive && ms < 8) ms = 8; // 独占下避免过小
            if (!exclusive && ms < 10) ms = 10;
            return ms;
        }

        // 事件优先创建（失败自动回退轮询或按用户强制）
        WasapiOut CreateOutWithPolicy(MMDevice dev, AudioClientShareMode mode, SyncModeOption syncPref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false;
            WasapiOut w;

            if (syncPref == SyncModeOption.Polling)
                return TryOut(dev, mode, false, bufMs, src);

            if (syncPref == SyncModeOption.Event)
            {
                w = TryOut(dev, mode, true, bufMs, src);
                if (w != null) { eventUsed = true; return w; }
                return TryOut(dev, mode, false, bufMs, src);
            }

            // Auto：事件优先
            w = TryOut(dev, mode, true, bufMs, src);
            if (w != null) { eventUsed = true; return w; }
            return TryOut(dev, mode, false, bufMs, src);
        }

        WasapiOut TryOut(MMDevice dev, AudioClientShareMode mode, bool eventSync, int bufMs, IWaveProvider src)
        {
            try
            {
                var w = new WasapiOut(dev, mode, eventSync, bufMs);
                w.Init(src);
                if (Logger.Enabled) Logger.Info($"WasapiOut OK: {dev.FriendlyName} | {(mode==AudioClientShareMode.Exclusive?"Ex":"Sh")} | {(eventSync?"Evt":"Poll")} | {bufMs}ms");
                return w;
            }
            catch (Exception ex)
            {
                if (Logger.Enabled) Logger.Info($"WasapiOut FAIL: {dev.FriendlyName} | {(mode==AudioClientShareMode.Exclusive?"Ex":"Sh")} | {(eventSync?"Evt":"Poll")} | {bufMs}ms | {ex.Message}");
                return null;
            }
        }

        static string Fmt(WaveFormat wf) => wf == null ? "-" : $"{wf.SampleRate}Hz/{wf.BitsPerSample}bit/{wf.Channels}ch";
        static bool FormatsEqual(WaveFormat a, WaveFormat b) => a != null && b != null && a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels;

        // 仅启动失败时的静默清理
        void CleanupCreated()
        {
            try { _capture?.Dispose(); } catch { } _capture = null;
            try { _mainOut?.Dispose(); } catch { } _mainOut = null;
            try { _auxOut?.Dispose();  } catch { } _auxOut  = null;
            try { _resMain?.Dispose(); } catch { } _resMain = null;
            try { _resAux?.Dispose();  } catch { } _resAux  = null;
            _bufMain = null; _bufAux = null; _srcMain = null; _srcAux = null;
        }

        public void Stop()
        {
            if (!_running)
            {
                CleanupCreated();
                return;
            }
            try { _capture?.StopRecording(); } catch { }
            try { _mainOut?.Stop(); } catch { }
            try { _auxOut ?.Stop(); } catch { }
            Thread.Sleep(25);
            CleanupCreated();
            _running = false;
        }

        // 状态提供给设置窗口
        public StatusSnapshot GetStatusSnapshot()
        {
            return new StatusSnapshot
            {
                Running = _running,
                InputRole = _inRole,
                InputFormat = _inFmt,
                InputDevice = _inName,
                MainDevice = _outMain?.FriendlyName ?? "-",
                AuxDevice  = _outAux ?.FriendlyName ?? "-",
                MainMode = _mainOut != null ? (_mainIsExclusive ? "独占" : "共享") : "-",
                AuxMode  = _auxOut  != null ? (_auxIsExclusive  ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEventUsed ? "事件" : "轮询") : "-",
                AuxSync  = _auxOut  != null ? (_auxEventUsed  ? "事件" : "轮询") : "-",
                MainFormat = _mainFmt, AuxFormat = _auxFmt,
                MainBufferMs = _mainOut != null ? _mainBufMsEff : 0,
                AuxBufferMs  = _auxOut  != null ? _auxBufMsEff  : 0
            };
        }

        // 工具
        MMDevice FindById(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            try
            {
                var col = _mm.EnumerateAudioEndPoints(flow, DeviceState.Active);
                foreach (var d in col) if (d.ID == id) return d;
            } catch { }
            return null;
        }

        void EnsureAutoStart(bool enable)
        {
            try
            {
                using (var run = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (run == null) return;
                    const string name = "MirrorAudio";
                    if (enable) run.SetValue(name, "\"" + Application.ExecutablePath + "\"");
                    else run.DeleteValue(name, false);
                }
            } catch { }
        }

        public void Dispose()
        {
            try { _mm?.UnregisterEndpointNotificationCallback(this); } catch { }
            Stop();
            _tray.Visible = false;
            _menu.Dispose();
            _tray.Dispose();
            _debounce.Dispose();
        }
    }
}
