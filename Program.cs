
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
using WinFormsTimer = System.Windows.Forms.Timer;

namespace MirrorAudio
{
    // -------- Logger --------
    static class Logger
    {
        public static bool Enabled;
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        static readonly string CrashPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
        public static void Info(string s) { if (!Enabled) return; try { File.AppendAllText(LogPath, "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + s + Environment.NewLine); } catch { } }
        public static void Crash(string where, Exception ex) { if (ex == null) return; try { File.AppendAllText(CrashPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}{Environment.NewLine}"); } catch { } }
    }

    // -------- Entry --------
    static class Program
    {
        static Mutex _mtx;
        [STAThread]
        static void Main()
        {
            bool ok;
            _mtx = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out ok);
            if (!ok) return;

            Application.ThreadException += (s, e) => Logger.Crash("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Crash("NonUI", e.ExceptionObject as Exception);

            try { MediaFoundationApi.Startup(); } catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) { Application.Run(); }

            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    // -------- Enums & Settings --------
    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    [DataContract]
    public enum InputFormatStrategy
    {
        [EnumMember] SystemMix = 0,
        [EnumMember] Specify24_48000,
        [EnumMember] Specify24_96000,
        [EnumMember] Specify24_192000,
        [EnumMember] Specify32f_48000,
        [EnumMember] Specify32f_96000,
        [EnumMember] Specify32f_192000,
        [EnumMember] Custom
    }

    [DataContract]
    public enum BufferAlignMode
    {
        [EnumMember] DefaultAlign,
        [EnumMember] MinAlign
    }

    [DataContract]
    public sealed class AppSettings
    {
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption  MainSync  = SyncModeOption.Auto,  AuxSync  = SyncModeOption.Auto;
        [DataMember] public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        [DataMember] public int AuxRate  =  48000, AuxBits  = 16, AuxBufMs  = 150;
        [DataMember] public bool AutoStart = false, EnableLogging = false;
        [DataMember] public BufferAlignMode MainBufMode = BufferAlignMode.DefaultAlign;
        [DataMember] public BufferAlignMode AuxBufMode  = BufferAlignMode.DefaultAlign;

        // Input strategy
        [DataMember] public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        [DataMember] public int InputCustomSampleRate = 96000;
        [DataMember] public int InputCustomBitDepth = 24;

        // Resampler control
        [DataMember] public int  MainResamplerQuality = 60; // 60/50/40/30
        [DataMember] public int  AuxResamplerQuality  = 30;
        [DataMember] public bool MainForceInternalResamplerInShared = false;
        [DataMember] public bool AuxForceInternalResamplerInShared  = false;
    }

    public sealed class StatusSnapshot
    {
        public bool Running;
        public int MainBufferRequestedMs, AuxBufferRequestedMs;
        public double MainAlignedMultiple, AuxAlignedMultiple;
        public double MainBufferMultiple,  AuxBufferMultiple;
        public string InputRole, InputFormat, InputDevice;
        public string InputRequested, InputAccepted, InputMix;
        public string MainDevice, AuxDevice, MainMode, AuxMode, MainSync, AuxSync, MainFormat, AuxFormat;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs, AuxDefaultPeriodMs, AuxMinimumPeriodMs;
        public bool MainNoSRC, AuxNoSRC, MainResampling, AuxResampling;
        public bool MainInternalResampler, AuxInternalResampler;
        public bool MainMultiSRC, AuxMultiSRC;
        public int MainInternalResamplerQuality, AuxInternalResamplerQuality;
    }

    static class Config
    {
        static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
        static readonly string FilePath = Path.Combine(Dir, "settings.json");
        public static AppSettings Load()
        {
            try {
                if (!File.Exists(FilePath)) return new AppSettings();
                using (var fs = File.OpenRead(FilePath))
                    return (AppSettings)new DataContractJsonSerializer(typeof(AppSettings)).ReadObject(fs);
            } catch { return new AppSettings(); }
        }
        public static void Save(AppSettings s)
        {
            try {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                using (var fs = File.Create(FilePath))
                    new DataContractJsonSerializer(typeof(AppSettings)).WriteObject(fs, s);
            } catch { }
        }
    }

    // -------- Tray App --------
    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly WinFormsTimer _debounce = new WinFormsTimer { Interval = 400 };

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture; BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux; WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;

        bool _running;
        bool _mainIsExclusive, _mainEventSyncUsed;
        bool _auxIsExclusive,  _auxEventSyncUsed;
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        string _inReqStr = "-", _inAccStr = "-", _inMixStr = "-";
        bool _mainNoSRC, _auxNoSRC, _mainResampling, _auxResampling;
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(4);

        public TrayApp()
        {
            Logger.Enabled = _cfg.EnableLogging;
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                if (File.Exists(icoPath)) _tray.Icon = new Icon(icoPath);
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }

            _tray.Visible = true;
            _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop  = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, (s, e) => OnSettings());
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s, e) => Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s, e) => { _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result; Logger.Enabled = _cfg.EnableLogging; Config.Save(_cfg);
                    EnsureAutoStart(_cfg.AutoStart); StartOrRestart();
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
                    if (enable) run.SetValue(name, """ + Application.ExecutablePath + """);
                    else run.DeleteValue(name, false);
                }
            }
            catch { }
        }

        // IMMNotificationClient
        public void OnDeviceStateChanged(string id, DeviceState st) { if (IsRelevant(id)) Debounce(); }
        public void OnDeviceAdded(string id) { if (IsRelevant(id)) Debounce(); }
        public void OnDeviceRemoved(string id) { if (IsRelevant(id)) Debounce(); }
        public void OnDefaultDeviceChanged(DataFlow f, Role r, string id) { if (f == DataFlow.Render && r == Role.Multimedia && string.IsNullOrEmpty(_cfg != null ? _cfg.InputDeviceId : null)) Debounce(); }
        public void OnPropertyValueChanged(string id, PropertyKey key) { if (IsRelevant(id)) Debounce(); }
        bool IsRelevant(string id)
        {
            if (string.IsNullOrEmpty(id) || _cfg == null) return false;
            return id.Equals(_cfg.InputDeviceId, StringComparison.OrdinalIgnoreCase)
                || id.Equals(_cfg.MainDeviceId, StringComparison.OrdinalIgnoreCase)
                || id.Equals(_cfg.AuxDeviceId, StringComparison.OrdinalIgnoreCase);
        }
        void Debounce() { _debounce.Stop(); _debounce.Start(); }

        // -------- Main pipeline --------
        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            _inDev = FindById(_cfg.InputDeviceId, DataFlow.Capture) ?? FindById(_cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux  = FindById(_cfg.AuxDeviceId,  DataFlow.Render);
            _inDevName = _inDev != null ? _inDev.FriendlyName : "-";
            if (_outMain == null || _outAux == null) { MessageBox.Show("请在“设置”选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            WaveFormat inFmt;
            WaveFormat inputRequested = null, inputAccepted = null, inputMix = null;

            if (_inDev.DataFlow == DataFlow.Capture)
            {
                _inRoleStr = "录音";
                try { inputMix = _inDev.AudioClient.MixFormat; } catch { }

                // We keep shared capture for stability
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                inFmt = cap.WaveFormat;
                _capture = cap;

                _inReqStr = Fmt(inFmt);
                _inAccStr = Fmt(inFmt);
                _inMixStr = Fmt(inputMix);
            }
            else
            {
                _inRoleStr = "环回";
                var cap = new WasapiLoopbackCapture(_inDev);
                try { inputMix = _inDev.AudioClient.MixFormat; } catch { }
                var req = new InputFormatRequest
                {
                    Strategy = _cfg.InputFormatStrategy,
                    CustomSampleRate = _cfg.InputCustomSampleRate,
                    CustomBitDepth = _cfg.InputCustomBitDepth,
                    Channels = 2
                };
                string log;
                WaveFormat accepted;
                var wf = InputFormatHelper.NegotiateLoopbackFormat(_inDev, req, out log, out inputMix, out accepted, out inputRequested);
                if (wf != null) cap.WaveFormat = wf;
                _capture = cap; inFmt = cap.WaveFormat;
                _inReqStr = InputFormatHelper.Fmt(inputRequested);
                _inAccStr = InputFormatHelper.Fmt(accepted ?? inFmt);
                _inMixStr = InputFormatHelper.Fmt(inputMix);
                if (Logger.Enabled) Logger.Info("Loopback negotiation:\n" + log);
            }

            _inFmtStr = Fmt(inFmt);
            if (Logger.Enabled) Logger.Info("Input: " + _inDev.FriendlyName + " | " + _inFmtStr + " | " + _inRoleStr);

            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 3.5, 60)) };
            _bufAux  = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs  * 4.0, 120)) };

            GetPeriods(_outMain, out _defMainMs, out _minMainMs);
            GetPeriods(_outAux,  out _defAuxMs,  out _minAuxMs);

            // ---- Main out ----
            _srcMain = _bufMain; _resMain = null; _mainIsExclusive = false; _mainEventSyncUsed = false; _mainBufEffectiveMs = _cfg.MainBufMs; _mainFmtStr = "-";
            _mainNoSRC = false; _mainResampling = false;
            var desiredMain = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
            bool isLoopMain = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outMain.ID;
            bool wantExMain = (_cfg.MainShare == ShareModeOption.Exclusive || _cfg.MainShare == ShareModeOption.Auto) && !isLoopMain;
            if (isLoopMain && _cfg.MainShare != ShareModeOption.Shared) MessageBox.Show("输入为主设备环回，独占冲突，主通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

            WaveFormat mainTargetFmt = null;

            if (wantExMain && SupportsExclusive(_outMain, desiredMain))
            {
                bool needRateChange = (inFmt.SampleRate != desiredMain.SampleRate) || (inFmt.Channels != desiredMain.Channels);
                if (needRateChange) _srcMain = _resMain = new MediaFoundationResampler(_bufMain, desiredMain) { ResamplerQuality = _cfg.MainResamplerQuality };
                int ms = BufAligned(_cfg.MainBufMs, true, _defMainMs, _minMainMs, _cfg.MainBufMode);
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                if (_mainOut != null) { _mainIsExclusive = true; _mainBufEffectiveMs = ms; _mainFmtStr = Fmt(desiredMain); mainTargetFmt = desiredMain; _mainResampling = needRateChange; _mainNoSRC = !needRateChange; }
            }
            if (_mainOut == null && _cfg.MainShare == ShareModeOption.Exclusive)
            {
                MessageBox.Show("主通道独占失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning); Cleanup(); return;
            }
            if (_mainOut == null)
            {
                int ms = BufAligned(_cfg.MainBufMs, false, _defMainMs, 0, _cfg.MainBufMode);
                // Force internal resampler in shared if requested
                try { mainTargetFmt = _outMain.AudioClient.MixFormat; } catch { mainTargetFmt = null; }
                IWaveProvider feed = _bufMain;
                if (_cfg.MainForceInternalResamplerInShared && mainTargetFmt != null)
                {
                    bool needChange = feed.WaveFormat.SampleRate != mainTargetFmt.SampleRate
                                   || feed.WaveFormat.Channels != mainTargetFmt.Channels
                                   || feed.WaveFormat.BitsPerSample != mainTargetFmt.BitsPerSample;
                    if (needChange)
                    {
                        _srcMain = _resMain = new MediaFoundationResampler(feed, mainTargetFmt) { ResamplerQuality = _cfg.MainResamplerQuality };
                        feed = _srcMain; _mainResampling = true; _mainNoSRC = false;
                    }
                }
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, ms, feed, out _mainEventSyncUsed);
                if (_mainOut == null) { MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); Cleanup(); return; }
                _mainBufEffectiveMs = ms;
                try { if (mainTargetFmt == null) mainTargetFmt = _outMain.AudioClient.MixFormat; _mainFmtStr = Fmt(mainTargetFmt); } catch { _mainFmtStr = "系统混音"; }
                _mainResampling = _mainResampling || (inFmt.SampleRate != (mainTargetFmt != null ? mainTargetFmt.SampleRate : inFmt.SampleRate)
                                                   || inFmt.Channels != (mainTargetFmt != null ? mainTargetFmt.Channels : inFmt.Channels));
                _mainNoSRC = !_mainResampling;
            }

            // ---- Aux out ----
            _srcAux = _bufAux; _resAux = null; _auxIsExclusive = false; _auxEventSyncUsed = false; _auxBufEffectiveMs = _cfg.AuxBufMs; _auxFmtStr = "-";
            _auxNoSRC = false; _auxResampling = false;
            var desiredAux = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);
            bool isLoopAux = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outAux.ID;
            bool wantExAux = (_cfg.AuxShare == ShareModeOption.Exclusive || _cfg.AuxShare == ShareModeOption.Auto) && !isLoopAux;
            if (isLoopAux && _cfg.AuxShare != ShareModeOption.Shared) MessageBox.Show("输入为副设备环回，独占冲突，副通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

            WaveFormat auxTargetFmt = null;

            if (wantExAux && SupportsExclusive(_outAux, desiredAux))
            {
                bool needRateChange = (inFmt.SampleRate != desiredAux.SampleRate) || (inFmt.Channels != desiredAux.Channels);
                if (needRateChange) _srcAux = _resAux = new MediaFoundationResampler(_bufAux, desiredAux) { ResamplerQuality = _cfg.AuxResamplerQuality };
                int ms = BufAligned(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs, _cfg.AuxBufMode);
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, ms, _srcAux, out _auxEventSyncUsed);
                if (_auxOut != null) { _auxIsExclusive = true; _auxBufEffectiveMs = ms; _auxFmtStr = Fmt(desiredAux); auxTargetFmt = desiredAux; _auxResampling = needRateChange; _auxNoSRC = !needRateChange; }
            }
            if (_auxOut == null && _cfg.AuxShare == ShareModeOption.Exclusive)
            {
                MessageBox.Show("副通道独占失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Warning); Cleanup(); return;
            }
            if (_auxOut == null)
            {
                int ms = BufAligned(_cfg.AuxBufMs, false, _defAuxMs, 0, _cfg.AuxBufMode);
                try { auxTargetFmt = _outAux.AudioClient.MixFormat; } catch { auxTargetFmt = null; }
                IWaveProvider feed = _bufAux;
                if (_cfg.AuxForceInternalResamplerInShared && auxTargetFmt != null)
                {
                    bool needChange = feed.WaveFormat.SampleRate != auxTargetFmt.SampleRate
                                   || feed.WaveFormat.Channels != auxTargetFmt.Channels
                                   || feed.WaveFormat.BitsPerSample != auxTargetFmt.BitsPerSample;
                    if (needChange)
                    {
                        _srcAux = _resAux = new MediaFoundationResampler(feed, auxTargetFmt) { ResamplerQuality = _cfg.AuxResamplerQuality };
                        feed = _srcAux; _auxResampling = true; _auxNoSRC = false;
                    }
                }
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, ms, feed, out _auxEventSyncUsed);
                if (_auxOut == null) { MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); Cleanup(); return; }
                _auxBufEffectiveMs = ms;
                try { if (auxTargetFmt == null) auxTargetFmt = _outAux.AudioClient.MixFormat; _auxFmtStr = Fmt(auxTargetFmt); } catch { _auxFmtStr = "系统混音"; }
                _auxResampling = _auxResampling || (inFmt.SampleRate != (auxTargetFmt != null ? auxTargetFmt.SampleRate : inFmt.SampleRate)
                                                 || inFmt.Channels != (auxTargetFmt != null ? auxTargetFmt.Channels : inFmt.Channels));
                _auxNoSRC = !_auxResampling;
            }

            _capture.DataAvailable += OnIn;
            _capture.RecordingStopped += OnStopRec;
            try
            {
                if (Logger.Enabled) Logger.Info("输入开始录制: " + _inDev.FriendlyName + " / " + inFmt);
                _capture.StartRecording();

                // prime buffers
                try
                {
                    int waited = 0;
                    int mainTarget = Math.Max((int)(_cfg.MainBufMs * 0.6), 20);
                    int auxTarget  = Math.Max((int)(_cfg.AuxBufMs  * 0.6), 40);
                    for (; waited < 300; waited += 5)
                    {
                        var m = _bufMain != null ? _bufMain.BufferedDuration.TotalMilliseconds : 0;
                        var a = _bufAux  != null ? _bufAux .BufferedDuration.TotalMilliseconds : 0;
                        if (m >= mainTarget && a >= auxTarget) break;
                        Thread.Sleep(5);
                    }
                }
                catch { }

                _mainOut.Play(); _auxOut.Play(); _running = true;
                if (Logger.Enabled)
                {
                    Logger.Info("Main: " + (_mainIsExclusive ? "独占" : "共享") + " | " + (_mainEventSyncUsed ? "事件" : "轮询") + " | " + _mainBufEffectiveMs + "ms");
                    Logger.Info("Aux : " + (_auxIsExclusive  ? "独占" : "共享") + " | " + (_auxEventSyncUsed  ? "事件" : "轮询") + " | " + _auxBufEffectiveMs  + "ms");
                }
            }
            catch (Exception ex) { Logger.Crash("Start", ex); MessageBox.Show("启动失败：" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); Stop(); }
        }

        void OnIn(object s, WaveInEventArgs e) { _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded); _bufAux.AddSamples(e.Buffer, 0, e.BytesRecorded); }
        void OnStopRec(object s, StoppedEventArgs e) { try { _bufMain?.ClearBuffer(); } catch { } try { _bufAux?.ClearBuffer(); } catch { } }

        public void Stop()
        {
            if (!_running) { DisposeAll(); return; }
            try { _capture?.StopRecording(); } catch { }
            try { _mainOut?.Stop(); } catch { }
            try { _auxOut?.Stop(); } catch { }
            Thread.Sleep(20); DisposeAll(); _running = false; _tray.ShowBalloonTip(600, "MirrorAudio", "已停止", ToolTipIcon.Info);
        }

        void DisposeAll()
        {
            try { if (_capture != null) { _capture.DataAvailable -= OnIn; _capture.RecordingStopped -= OnStopRec; _capture.Dispose(); } } catch { } _capture = null;
            try { _mainOut?.Dispose(); } catch { } _mainOut = null;
            try { _auxOut ?.Dispose(); } catch { } _auxOut  = null;
            try { _resMain?.Dispose(); } catch { } _resMain = null;
            try { _resAux ?.Dispose(); } catch { } _resAux  = null;
            _bufMain = null; _bufAux = null;
        }
        void Cleanup() { DisposeAll(); }

        // -------- Status to SettingsForm --------
        public StatusSnapshot GetStatusSnapshot()
        {
            // buffer-alignment multiple
            double mainMul = 0, auxMul = 0;
            try
            {
                if (_mainBufEffectiveMs > 0)
                {
                    double baseMs = (_mainIsExclusive && _minMainMs > 0) ? _minMainMs : _defMainMs;
                    if (baseMs > 0) mainMul = Math.Round(_mainBufEffectiveMs / baseMs, 2);
                }
                if (_auxBufEffectiveMs > 0)
                {
                    double baseMs = (_auxIsExclusive && _minAuxMs > 0) ? _minAuxMs : _defAuxMs;
                    if (baseMs > 0) auxMul = Math.Round(_auxBufEffectiveMs / baseMs, 2);
                }
            }
            catch { }

            // internal & multi SRC
            bool mainInternal = _resMain != null;
            bool auxInternal  = _resAux  != null;
            bool mainMulti = false, auxMulti = false;
            try
            {
                if (!_mainIsExclusive && mainInternal)
                {
                    WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }
                    if (mix != null && _resMain != null)
                    {
                        var f = _resMain.WaveFormat;
                        if (f.SampleRate != mix.SampleRate || f.Channels != mix.Channels || f.BitsPerSample != mix.BitsPerSample) mainMulti = true;
                    }
                }
                if (!_auxIsExclusive && auxInternal)
                {
                    WaveFormat mix = null; try { mix = _outAux.AudioClient.MixFormat; } catch { }
                    if (mix != null && _resAux != null)
                    {
                        var f = _resAux.WaveFormat;
                        if (f.SampleRate != mix.SampleRate || f.Channels != mix.Channels || f.BitsPerSample != mix.BitsPerSample) auxMulti = true;
                    }
                }
            }
            catch { }

            return new StatusSnapshot
            {
                Running = _running,
                InputRole = _inRoleStr, InputFormat = _inFmtStr, InputDevice = _inDevName,
                InputRequested = _inReqStr, InputAccepted = _inAccStr, InputMix = _inMixStr,
                MainDevice = _outMain != null ? _outMain.FriendlyName : SafeName(_cfg.MainDeviceId, DataFlow.Render),
                AuxDevice  = _outAux  != null ? _outAux .FriendlyName : SafeName(_cfg.AuxDeviceId , DataFlow.Render),
                MainMode = _mainOut != null ? (_mainIsExclusive ? "独占" : "共享") : "-",
                AuxMode  = _auxOut  != null ? (_auxIsExclusive  ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEventSyncUsed ? "事件" : "轮询") : "-",
                AuxSync  = _auxOut  != null ? (_auxEventSyncUsed  ? "事件" : "轮询") : "-",
                MainFormat = _mainOut != null ? _mainFmtStr : "-",
                AuxFormat  = _auxOut  != null ? _auxFmtStr  : "-",
                MainBufferRequestedMs = _cfg.MainBufMs, AuxBufferRequestedMs = _cfg.AuxBufMs,
                MainBufferMs = _mainOut != null ? _mainBufEffectiveMs : 0, AuxBufferMs = _auxOut != null ? _auxBufEffectiveMs : 0,
                MainDefaultPeriodMs = _defMainMs, MainMinimumPeriodMs = _minMainMs,
                AuxDefaultPeriodMs  = _defAuxMs,  AuxMinimumPeriodMs  = _minAuxMs,
                MainAlignedMultiple = mainMul, AuxAlignedMultiple = auxMul,
                MainNoSRC = _mainNoSRC, AuxNoSRC = _auxNoSRC, MainResampling = _mainResampling, AuxResampling = _auxResampling,
                MainBufferMultiple = (_mainBufEffectiveMs > 0 && _minMainMs > 0) ? _mainBufEffectiveMs / _minMainMs : 0,
                AuxBufferMultiple  = (_auxBufEffectiveMs  > 0 && _minAuxMs  > 0) ? _auxBufEffectiveMs  / _minAuxMs  : 0,
                MainInternalResampler = mainInternal, AuxInternalResampler = auxInternal,
                MainMultiSRC = mainMulti, AuxMultiSRC = auxMulti,
                MainInternalResamplerQuality = _resMain != null ? _resMain.ResamplerQuality : 0,
                AuxInternalResamplerQuality  = _resAux  != null ? _resAux .ResamplerQuality  : 0
            };
        }

        string SafeName(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return "-";
            try { foreach (var d in _mm.EnumerateAudioEndPoints(flow, DeviceState.Active)) if (d.ID == id) return d.FriendlyName; } catch { }
            return "-";
        }
        MMDevice FindById(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            try { foreach (var d in _mm.EnumerateAudioEndPoints(flow, DeviceState.Active)) if (d.ID == id) return d; } catch { }
            return null;
        }
        static string Fmt(WaveFormat wf) { return wf == null ? "-" : (wf.SampleRate + "Hz/" + wf.BitsPerSample + "bit/" + wf.Channels + "ch"); }

        void GetPeriods(MMDevice dev, out double defMs, out double minMs)
        {
            defMs = 10; minMs = 2; if (dev == null) return; var id = dev.ID; Tuple<double, double> t;
            if (_periodCache.TryGetValue(id, out t)) { defMs = t.Item1; minMs = t.Item2; return; }
            try
            {
                long d100 = 0, m100 = 0; var ac = dev.AudioClient;
                var pD = ac.GetType().GetProperty("DefaultDevicePeriod");
                var pM = ac.GetType().GetProperty("MinimumDevicePeriod");
                if (pD != null) { var v = pD.GetValue(ac, null); if (v != null) d100 = Convert.ToInt64(v); }
                if (pM != null) { var v = pM.GetValue(ac, null); if (v != null) m100 = Convert.ToInt64(v); }
                if (d100 > 0) defMs = d100 / 10000.0; if (m100 > 0) minMs = m100 / 10000.0;
            }
            catch { }
            _periodCache[id] = Tuple.Create(defMs, minMs);
        }

        static bool SupportsExclusive(MMDevice d, WaveFormat f) { try { return d.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, f); } catch { return false; } }

        static int BufAligned(int wantMs, bool exclusive, double defMs, double minMs, BufferAlignMode mode)
        {
            if (exclusive)
            {
                double stepMin = (minMs > 0 ? minMs : (defMs > 0 ? defMs : 10.0));
                double stepDef = (defMs > 0 ? defMs : stepMin);
                int ms;
                if (mode == BufferAlignMode.MinAlign)
                {
                    ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepMin) * stepMin);
                    double floor = stepMin * 3.0;
                    if (ms < floor) ms = (int)Math.Ceiling(Math.Ceiling(floor / stepMin) * stepMin);
                    return ms;
                }
                else
                {
                    ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepDef) * stepDef);
                    double floor = stepDef * 3.0;
                    if (ms < floor) ms = (int)Math.Ceiling(Math.Ceiling(floor / stepDef) * stepDef);
                    return ms;
                }
            }
            else
            {
                double stepMin = (minMs > 0 ? minMs : (defMs > 0 ? defMs : 10.0));
                double stepDef = (defMs > 0 ? defMs : stepMin);
                int ms;
                if (mode == BufferAlignMode.MinAlign)
                {
                    ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepMin) * stepMin);
                }
                else
                {
                    ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepDef) * stepDef);
                }
                double floor = (defMs > 0 ? defMs : stepMin) * 2.0;
                if (ms < floor)
                {
                    double step = (mode == BufferAlignMode.MinAlign ? stepMin : stepDef);
                    ms = (int)Math.Ceiling(Math.Ceiling(floor / step) * step);
                }
                return ms;
            }
        }

        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption pref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false; WasapiOut w;
            if (pref == SyncModeOption.Polling) return TryOut(dev, mode, false, bufMs, src);
            if (pref == SyncModeOption.Event) { w = TryOut(dev, mode, true, bufMs, src); if (w != null) { eventUsed = true; return w; } return TryOut(dev, mode, false, bufMs, src); }
            w = TryOut(dev, mode, true, bufMs, src); if (w != null) { eventUsed = true; return w; } return TryOut(dev, mode, false, bufMs, src);
        }
        WasapiOut TryOut(MMDevice dev, AudioClientShareMode mode, bool ev, int ms, IWaveProvider src)
        {
            try { var w = new WasapiOut(dev, mode, ev, ms); w.Init(src); if (Logger.Enabled) Logger.Info($"OK {dev.FriendlyName} | {mode} | {(ev ? "event" : "poll")} | {ms}ms"); return w; }
            catch (Exception ex) { if (Logger.Enabled) Logger.Info($"Fail {dev.FriendlyName} | {mode} | {(ev ? "event" : "poll")} | {ms}ms | 0x{((uint)ex.HResult):X8} {ex.Message}"); return null; }
        }

        public void Dispose() { Stop(); try { _mm?.UnregisterEndpointNotificationCallback(this); } catch { } _debounce.Dispose(); _tray.Visible = false; _tray.Dispose(); _menu.Dispose(); }
    }

    // -------- Input negotiation helpers --------
    public sealed class InputFormatRequest
    {
        public InputFormatStrategy Strategy = InputFormatStrategy.SystemMix;
        public int CustomSampleRate = 48000;
        public int CustomBitDepth = 24;
        public int Channels = 2;
    }

    public static class InputFormatHelper
    {
        public static WaveFormat BuildWaveFormat(InputFormatStrategy strategy, int customRate, int customBits, int channels)
        {
            switch (strategy)
            {
                case InputFormatStrategy.SystemMix: return null;
                case InputFormatStrategy.Specify24_48000:  return CreatePcm24(48000, channels);
                case InputFormatStrategy.Specify24_96000:  return CreatePcm24(96000, channels);
                case InputFormatStrategy.Specify24_192000: return CreatePcm24(192000, channels);
                case InputFormatStrategy.Specify32f_48000:  return WaveFormat.CreateIeeeFloatWaveFormat(48000, channels);
                case InputFormatStrategy.Specify32f_96000:  return WaveFormat.CreateIeeeFloatWaveFormat(96000, channels);
                case InputFormatStrategy.Specify32f_192000: return WaveFormat.CreateIeeeFloatWaveFormat(192000, channels);
                case InputFormatStrategy.Custom:
                    if (customBits >= 32) return WaveFormat.CreateIeeeFloatWaveFormat(customRate, channels);
                    if (customBits == 24) return CreatePcm24(customRate, channels);
                    return new WaveFormat(customRate, customBits, channels);
                default: return null;
            }
        }

        public static WaveFormat CreatePcm24(int sampleRate, int channels)
        {
            return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Extensible, sampleRate, channels, sampleRate * channels * 3, 3, 24);
        }

        public static WaveFormat NegotiateLoopbackFormat(MMDevice device, InputFormatRequest request,
            out string log, out WaveFormat mixFormat, out WaveFormat acceptedFormat, out WaveFormat requestedFormat)
        {
            var sb = new System.Text.StringBuilder();
            mixFormat = null; acceptedFormat = null; requestedFormat = null;
            try { mixFormat = device.AudioClient.MixFormat; } catch { }

            var desired = BuildWaveFormat(request.Strategy, request.CustomSampleRate, request.CustomBitDepth, request.Channels);
            requestedFormat = desired;
            if (mixFormat != null) sb.AppendLine("Device Mix: " + Fmt(mixFormat));
            if (desired == null)
            {
                sb.AppendLine("Request: SystemMix (use engine-provided mix).");
                acceptedFormat = mixFormat;
                log = sb.ToString();
                return null;
            }

            WaveFormatExtensible closest = null;
            bool ok = false;
            try { ok = device.AudioClient.IsFormatSupported(AudioClientShareMode.Shared, desired, out closest); }
            catch { ok = false; }

            sb.AppendLine("Request: " + Fmt(desired) + " -> Supported: " + (ok ? "Yes" : "No"));
            if (!ok && closest != null) sb.AppendLine("Closest: " + Fmt(closest));

            if (ok) { acceptedFormat = desired; log = sb.ToString(); return desired; }
            if (closest != null) { acceptedFormat = closest; log = sb.ToString(); return closest; }

            // fallback to mix
            acceptedFormat = mixFormat;
            log = sb.ToString();
            return null;
        }

        public static string Fmt(WaveFormat wf) => wf == null ? "-" : (wf.SampleRate + "Hz/" + wf.BitsPerSample + "bit/" + wf.Channels + "ch");
    }
}
