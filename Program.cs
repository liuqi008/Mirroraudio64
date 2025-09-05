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

namespace MirrorAudio
{
    // 轻量日志：先判定再打印，避免无谓字符串分配
    static class Logger
    {
        public static bool Enabled;
        static readonly string _logPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        public static void Info(string msg)
        {
            if (!Enabled) return;
            try { File.AppendAllText(_logPath, "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + "\r\n"); } catch { }
        }
        public static void Crash(string where, Exception ex)
        {
            if (ex == null) return;
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
                File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n");
            } catch { }
        }
    }

    static class Program
    {
        static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            bool createdNew;
            _mutex = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out createdNew);
            if (!createdNew) return;

            Application.ThreadException += (s, e) => Logger.Crash("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Crash("Non-UI", e.ExceptionObject as Exception);

            try { MediaFoundationApi.Startup(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) Application.Run();
            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    // —— 配置 —— //
    [DataContract]
    public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract]
    public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    [DataContract]
    public class AppSettings
    {
        [DataMember] public string InputDeviceId;
        [DataMember] public string MainDeviceId;
        [DataMember] public string AuxDeviceId;

        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto;
        [DataMember] public SyncModeOption  MainSync  = SyncModeOption.Auto;
        [DataMember] public int MainRate  = 192000;
        [DataMember] public int MainBits  = 24;
        [DataMember] public int MainBufMs = 12;

        [DataMember] public ShareModeOption AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption  AuxSync  = SyncModeOption.Auto;
        [DataMember] public int AuxRate  = 48000;   // 仅独占生效
        [DataMember] public int AuxBits  = 16;      // 仅独占生效
        [DataMember] public int AuxBufMs = 150;

        [DataMember] public bool AutoStart = false;
        [DataMember] public bool EnableLogging = false;
    }

    public class StatusSnapshot
    {
        public bool Running;
        public string InputRole;
        public string InputFormat;
        public string InputDevice;
        public string MainDevice;
        public string AuxDevice;
        public string MainMode;
        public string AuxMode;
        public string MainSync;
        public string AuxSync;
        public string MainFormat;
        public string AuxFormat;
        public int MainBufferMs;
        public int AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs;
        public double AuxDefaultPeriodMs,  AuxMinimumPeriodMs;
    }

    static class Config
    {
        static readonly string Dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
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
            } catch { return new AppSettings(); }
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
            } catch { }
        }
    }

    // —— 主程序 —— //
    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm;               // 单实例，复用
        readonly System.Windows.Forms.Timer _debounce;             // 单实例，复用
        bool _running;

        // 设备 & 音频对象
        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture;
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux;
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;

        // 最终模式/格式/周期
        bool _mainIsExclusive, _mainEventSyncUsed, _auxIsExclusive, _auxEventSyncUsed;
        int  _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-";
        string _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs, _minMainMs, _defAuxMs, _minAuxMs;

        // 设备周期缓存（降低反射与COM调用）
        readonly Dictionary<string, Tuple<double,double>> _periodCache = new Dictionary<string, Tuple<double,double>>(4);

        public TrayApp()
        {
            // 日志开关
            Logger.Enabled = _cfg.EnableLogging;

            _mm = new MMDeviceEnumerator();
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            // Tray icon
            try
            {
                if (File.Exists("MirrorAudio.ico")) _tray.Icon = new Icon("MirrorAudio.ico");
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }

            _tray.Visible = true;
            _tray.Text = "MirrorAudio";

            // 精简托盘菜单
            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, OnStartClick);
            var miStop  = new ToolStripMenuItem("停止(&T)", null, OnStopClick);
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, OnSettingsClick);
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s,e)=> Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)", null, (s,e)=> { Stop(); Application.Exit(); });

            _menu.Items.AddRange(new ToolStripItem[]{ miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce = new System.Windows.Forms.Timer();
            _debounce.Interval = 400;
            _debounce.Tick += OnDebounceTick;

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        void OnStartClick(object sender, EventArgs e) => StartOrRestart();
        void OnStopClick(object sender, EventArgs e)  => Stop();
        void OnSettingsClick(object sender, EventArgs e)
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

        // —— IMMNotificationClient（事件驱动热插拔自愈） —— //
        public void OnDeviceStateChanged(string id, DeviceState st) { if (IsRelevant(id)) DebounceRestart(); }
        public void OnDeviceAdded(string id)  { if (IsRelevant(id)) DebounceRestart(); }
        public void OnDeviceRemoved(string id){ if (IsRelevant(id)) DebounceRestart(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string id)
        {
            if (flow == DataFlow.Render && role == Role.Multimedia && string.IsNullOrEmpty(_cfg != null ? _cfg.InputDeviceId : null))
                DebounceRestart();
        }
        public void OnPropertyValueChanged(string id, PropertyKey key) { if (IsRelevant(id)) DebounceRestart(); }

        bool IsRelevant(string id)
        {
            if (string.IsNullOrEmpty(id) || _cfg == null) return false;
            return string.Equals(id, _cfg.InputDeviceId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.MainDeviceId,  StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.AuxDeviceId,   StringComparison.OrdinalIgnoreCase);
        }

        void DebounceRestart()
        {
            _debounce.Stop();
            _debounce.Start();
        }
        void OnDebounceTick(object sender, EventArgs e)
        {
            _debounce.Stop();
            StartOrRestart();
        }

        // —— 主流程 —— //
        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            // 输入优先取配置，否则默认渲染环回
            _inDev   = FindById(_cfg.InputDeviceId, DataFlow.Capture) ?? FindById(_cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux  = FindById(_cfg.AuxDeviceId,  DataFlow.Render);

            _inDevName = _inDev != null ? _inDev.FriendlyName : "-";
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
                _capture = cap; inFmt = cap.WaveFormat; _inRoleStr = "录音";
            }
            else
            {
                var cap = new WasapiLoopbackCapture(_inDev);
                _capture = cap; inFmt = cap.WaveFormat; _inRoleStr = "环回";
            }
            _inFmtStr = Fmt(inFmt);
            if (Logger.Enabled) Logger.Info("Input: " + _inDev.FriendlyName + " | " + _inFmtStr + " | " + _inRoleStr);

            // 两路桥接缓冲（确保足够大避免频繁扩容）
            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 8, 120)) };
            _bufAux  = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs  * 6, 150)) };

            // 设备周期（带缓存）
            GetDevicePeriodsMsCached(_outMain, out _defMainMs, out _minMainMs);
            GetDevicePeriodsMsCached(_outAux,  out _defAuxMs,  out _minAuxMs);

            // —— 主通道：独占/共享 + 事件/轮询 —— //
            _srcMain = _bufMain; _resMain = null;
            _mainIsExclusive = false; _mainEventSyncUsed = false; _mainBufEffectiveMs = _cfg.MainBufMs; _mainFmtStr = "-";

            var desiredMain = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
            bool loopbackConflictMain = (_inDev.DataFlow == DataFlow.Render) && string.Equals(_inDev.ID, _outMain.ID, StringComparison.OrdinalIgnoreCase);
            bool wantExclusiveMain = (_cfg.MainShare == ShareModeOption.Exclusive || _cfg.MainShare == ShareModeOption.Auto) && !loopbackConflictMain;

            if (loopbackConflictMain && (_cfg.MainShare != ShareModeOption.Shared))
                MessageBox.Show("通道1为主设备的环回，独占与环回冲突，主通道改走“共享”。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (wantExclusiveMain)
            {
                if (IsFormatSupportedExclusive(_outMain, desiredMain))
                {
                    if (!FormatsEqual(inFmt, desiredMain))
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, desiredMain) { ResamplerQuality = 50 };
                        _srcMain = _resMain;
                    }
                    int msEx = SafeBuf(_cfg.MainBufMs, true, _defMainMs, _minMainMs);
                    _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, msEx, _srcMain, out _mainEventSyncUsed);
                    if (_mainOut != null) { _mainIsExclusive = true; _mainBufEffectiveMs = msEx; _mainFmtStr = Fmt(desiredMain); }
                }
                if (_mainOut == null && _cfg.MainBits == 24)
                {
                    var fmt32 = new WaveFormat(_cfg.MainRate, 32, 2);
                    if (IsFormatSupportedExclusive(_outMain, fmt32))
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, fmt32) { ResamplerQuality = 50 };
                        _srcMain = _resMain;
                        int msEx2 = SafeBuf(_cfg.MainBufMs, true, _defMainMs, _minMainMs);
                        _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, msEx2, _srcMain, out _mainEventSyncUsed);
                        if (_mainOut != null) { _mainIsExclusive = true; _mainBufEffectiveMs = msEx2; _mainFmtStr = Fmt(fmt32); }
                    }
                }
                if (_mainOut == null && _cfg.MainShare == ShareModeOption.Exclusive)
                {
                    MessageBox.Show("主通道“强制独占”失败：请检查占用或格式。可尝试 32-bit 容器。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CleanupCreated(); return;
                }
            }

            if (_mainOut == null)
            {
                _srcMain = _bufMain;
                int msSh = SafeBuf(_cfg.MainBufMs, false, _defMainMs);
                _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, msSh, _srcMain, out _mainEventSyncUsed);
                if (_mainOut == null) { MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); CleanupCreated(); return; }
                _mainIsExclusive = false; _mainBufEffectiveMs = msSh;
                try { _mainFmtStr = Fmt(_outMain.AudioClient.MixFormat); } catch { _mainFmtStr = "系统混音"; }
            }

            // —— 副通道 —— //
            _srcAux = _bufAux; _resAux = null;
            _auxIsExclusive = false; _auxEventSyncUsed = false; _auxBufEffectiveMs = _cfg.AuxBufMs; _auxFmtStr = "-";
            var desiredAux = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);

            bool loopbackConflictAux = (_inDev.DataFlow == DataFlow.Render) && string.Equals(_inDev.ID, _outAux.ID, StringComparison.OrdinalIgnoreCase);
            bool wantExclusiveAux = (_cfg.AuxShare == ShareModeOption.Exclusive || _cfg.AuxShare == ShareModeOption.Auto) && !loopbackConflictAux;

            if (loopbackConflictAux && (_cfg.AuxShare != ShareModeOption.Shared))
                MessageBox.Show("通道1为副设备的环回，独占与环回冲突，副通道改走“共享”。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (wantExclusiveAux)
            {
                if (IsFormatSupportedExclusive(_outAux, desiredAux))
                {
                    if (!FormatsEqual(inFmt, desiredAux))
                    {
                        _resAux = new MediaFoundationResampler(_bufAux, desiredAux) { ResamplerQuality = 40 };
                        _srcAux = _resAux;
                    }
                    int msExA = SafeBuf(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs);
                    _auxOut = CreateOutWithPolicy(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, msExA, _srcAux, out _auxEventSyncUsed);
                    if (_auxOut != null) { _auxIsExclusive = true; _auxBufEffectiveMs = msExA; _auxFmtStr = Fmt(desiredAux); }
                }
                if (_auxOut == null && _cfg.AuxBits == 24)
                {
                    var fmt32A = new WaveFormat(_cfg.AuxRate, 32, 2);
                    if (IsFormatSupportedExclusive(_outAux, fmt32A))
                    {
                        _resAux = new MediaFoundationResampler(_bufAux, fmt32A) { ResamplerQuality = 40 };
                        _srcAux = _resAux;
                        int msExA2 = SafeBuf(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs);
                        _auxOut = CreateOutWithPolicy(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, msExA2, _srcAux, out _auxEventSyncUsed);
                        if (_auxOut != null) { _auxIsExclusive = true; _auxBufEffectiveMs = msExA2; _auxFmtStr = Fmt(fmt32A); }
                    }
                }
                if (_auxOut == null && _cfg.AuxShare == ShareModeOption.Exclusive)
                {
                    MessageBox.Show("副通道“强制独占”失败：请检查占用或格式。可尝试 32-bit 容器。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CleanupCreated(); return;
                }
            }

            if (_auxOut == null)
            {
                _srcAux = _bufAux;
                int msShA = SafeBuf(_cfg.AuxBufMs, false, _defAuxMs);
                _auxOut = CreateOutWithPolicy(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, msShA, _srcAux, out _auxEventSyncUsed);
                if (_auxOut == null) { MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); CleanupCreated(); return; }
                _auxIsExclusive = false; _auxBufEffectiveMs = msShA;
                try { _auxFmtStr = Fmt(_outAux.AudioClient.MixFormat); } catch { _auxFmtStr = "系统混音"; }
            }

            // —— 绑定与启动（方法组，避免闭包分配） —— //
            _capture.DataAvailable += OnCaptureDataAvailable;
            _capture.RecordingStopped += OnRecordingStopped;

            try
            {
                _mainOut.Play();
                _auxOut.Play();
                _capture.StartRecording();
                _running = true;

                if (Logger.Enabled)
                {
                    Logger.Info("Final Main: " + (_mainIsExclusive ? "独占" : "共享") + " | " + (_mainEventSyncUsed ? "事件" : "轮询") + " | buf=" + _mainBufEffectiveMs + "ms");
                    Logger.Info("Final Aux : " + (_auxIsExclusive  ? "独占" : "共享") + " | " + (_auxEventSyncUsed  ? "事件" : "轮询") + " | buf=" + _auxBufEffectiveMs + "ms");
                }
            }
            catch (Exception ex)
            {
                Logger.Crash("Start", ex);
                MessageBox.Show("启动失败：" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void OnCaptureDataAvailable(object sender, WaveInEventArgs e)
        {
            var buf = e.Buffer; int len = e.BytesRecorded;
            _bufMain.AddSamples(buf, 0, len);
            _bufAux.AddSamples(buf, 0, len);
        }
        void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (_bufMain != null) _bufMain.ClearBuffer();
            if (_bufAux  != null) _bufAux .ClearBuffer();
        }

        public void Stop()
        {
            if (!_running)
            {
                DisposeAll();
                return;
            }
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut.Stop();  } catch { }
            Thread.Sleep(30);
            DisposeAll();
            _running = false;
            _tray.ShowBalloonTip(800, "MirrorAudio", "已停止", ToolTipIcon.Info);
        }

        void DisposeAll()
        {
            try { if (_capture != null) { _capture.DataAvailable -= OnCaptureDataAvailable; _capture.RecordingStopped -= OnRecordingStopped; _capture.Dispose(); } } catch { } _capture = null;
            try { if (_mainOut != null) _mainOut.Dispose(); } catch { } _mainOut = null;
            try { if (_auxOut  != null) _auxOut .Dispose(); } catch { } _auxOut  = null;
            try { if (_resMain != null) _resMain.Dispose(); } catch { } _resMain = null;
            try { if (_resAux  != null) _resAux .Dispose(); } catch { } _resAux  = null;
            _bufMain = null; _bufAux = null;
        }
        void CleanupCreated()
        {
            DisposeAll();
        }
        // —— 即时状态供设置窗体按需读取 —— //
        public StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot
            {
                Running = _running,
                InputRole = _inRoleStr,
                InputFormat = _inFmtStr,
                InputDevice = _inDevName,
                MainDevice = _outMain != null ? _outMain.FriendlyName : SafeNameForId(_cfg.MainDeviceId, DataFlow.Render),
                AuxDevice  = _outAux  != null ? _outAux .FriendlyName : SafeNameForId(_cfg.AuxDeviceId,  DataFlow.Render),
                MainMode = _mainOut != null ? (_mainIsExclusive ? "独占" : "共享") : "-",
                AuxMode  = _auxOut  != null ? (_auxIsExclusive  ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEventSyncUsed ? "事件" : "轮询") : "-",
                AuxSync  = _auxOut  != null ? (_auxEventSyncUsed  ? "事件" : "轮询") : "-",
                MainFormat = _mainOut != null ? _mainFmtStr : "-",
                AuxFormat  = _auxOut  != null ? _auxFmtStr  : "-",
                MainBufferMs = _mainOut != null ? _mainBufEffectiveMs : 0,
                AuxBufferMs  = _auxOut  != null ? _auxBufEffectiveMs  : 0,
                MainDefaultPeriodMs = _defMainMs,
                MainMinimumPeriodMs = _minMainMs,
                AuxDefaultPeriodMs  = _defAuxMs,
                AuxMinimumPeriodMs  = _minAuxMs
            };
            return s;
        }

        string SafeNameForId(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return "-";
            try
            {
                foreach (var d in _mm.EnumerateAudioEndPoints(flow, DeviceState.Active))
                    if (d.ID == id) return d.FriendlyName;
            } catch { }
            return "-";
        }

        // —— 工具函数（去掉热路径 LINQ） —— //
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

        static string Fmt(WaveFormat wf)
        {
            if (wf == null) return "-";
            return wf.SampleRate + "Hz/" + wf.BitsPerSample + "bit/" + wf.Channels + "ch";
        }

        // 设备周期缓存：先查缓存，再反射获取一次
        void GetDevicePeriodsMsCached(MMDevice dev, out double defMs, out double minMs)
        {
            defMs = 10.0; minMs = 2.0;
            if (dev == null) return;
            var id = dev.ID;
            Tuple<double,double> t;
            if (_periodCache.TryGetValue(id, out t))
            {
                defMs = t.Item1; minMs = t.Item2; return;
            }

            try
            {
                long def100 = 0, min100 = 0;
                var ac = dev.AudioClient;
                var pDef = ac.GetType().GetProperty("DefaultDevicePeriod");
                var pMin = ac.GetType().GetProperty("MinimumDevicePeriod");
                if (pDef != null) { var v = pDef.GetValue(ac, null); if (v != null) def100 = Convert.ToInt64(v); }
                if (pMin != null) { var v = pMin.GetValue(ac, null); if (v != null) min100 = Convert.ToInt64(v); }

                if (def100 == 0 || min100 == 0)
                {
                    var m = ac.GetType().GetMethod("GetDevicePeriod");
                    if (m != null)
                    {
                        object[] args = new object[] { 0L, 0L };
                        m.Invoke(ac, args);
                        def100 = (long)args[0];
                        min100 = (long)args[1];
                    }
                }
                if (def100 > 0) defMs = def100 / 10000.0;
                if (min100 > 0) minMs = min100 / 10000.0;
            }
            catch { }

            _periodCache[id] = Tuple.Create(defMs, minMs);
        }

        static bool IsFormatSupportedExclusive(MMDevice dev, WaveFormat fmt)
        {
            try { return dev.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, fmt); }
            catch { return false; }
        }

        static bool FormatsEqual(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels;
        }

        static int SafeBuf(int desiredMs, bool exclusive, double defaultPeriodMs, double minPeriodMs = 0)
        {
            int ms = desiredMs;
            if (exclusive)
            {
                int floor = (int)Math.Ceiling(defaultPeriodMs * 3.0);
                if (ms < floor) ms = floor;
                if (minPeriodMs > 0)
                {
                    double k = Math.Ceiling(ms / minPeriodMs);
                    ms = (int)Math.Ceiling(k * minPeriodMs);
                }
            }
            else
            {
                int floor = (int)Math.Ceiling(defaultPeriodMs * 2.0);
                if (ms < floor) ms = floor;
            }
            return ms;
        }

        WasapiOut CreateOutWithPolicy(MMDevice dev, AudioClientShareMode mode, SyncModeOption syncPref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false;
            WasapiOut w;

            if (syncPref == SyncModeOption.Polling)
            {   // 强制轮询
                w = TryOut(dev, mode, false, bufMs, src);
                return w;
            }

            if (syncPref == SyncModeOption.Event)
            {   // 强制事件，失败回退轮询
                w = TryOut(dev, mode, true, bufMs, src);
                if (w != null) { eventUsed = true; return w; }
                w = TryOut(dev, mode, false, bufMs, src);
                return w;
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
                if (Logger.Enabled) Logger.Info($"WasapiOut OK: {dev.FriendlyName} | mode={mode} event={eventSync} buf={bufMs}ms");
                return w;
            }
            catch (Exception ex)
            {
                if (Logger.Enabled) Logger.Info($"WasapiOut failed: {dev.FriendlyName} | mode={mode} event={eventSync} buf={bufMs}ms | 0x{((uint)ex.HResult):X8} {ex.Message}");
                return null;
            }
        }

        public void Dispose()
        {
            Stop();
            try { if (_mm != null) _mm.UnregisterEndpointNotificationCallback(this); } catch { }
            _debounce.Dispose();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
        }
    }
}
