
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace MirrorAudio
{
    static class Program
    {
        static Mutex _mtx;

        [STAThread]
        static void Main()
        {
            bool ok;
            _mtx = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out ok);
            if (!ok) return;

            Application.ThreadException += (s, e) => { try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log"), e.Exception.ToString()); } catch { } };
            AppDomain.CurrentDomain.UnhandledException += (s, e) => { try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log"), (e.ExceptionObject as Exception)?.ToString()); } catch { } };

            try { MediaFoundationApi.Startup(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp())
            {
                Application.Run();
            }
            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }
    [DataContract] public enum BufferAlignMode { [EnumMember] DefaultAlign, [EnumMember] MinAlign }
    [DataContract] public enum InputFormatStrategy
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
    public sealed class AppSettings
    {
        [DataMember] public bool InputExclusive = false;
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption MainSync = SyncModeOption.Auto, AuxSync = SyncModeOption.Auto;
        [DataMember] public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        [DataMember] public int AuxRate = 48000, AuxBits = 16, AuxBufMs = 150;
        [DataMember] public BufferAlignMode MainBufMode = BufferAlignMode.DefaultAlign;
        [DataMember] public BufferAlignMode AuxBufMode = BufferAlignMode.DefaultAlign;
        [DataMember] public bool AutoStart = false, EnableLogging = false;
        [DataMember] public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        [DataMember] public int InputCustomSampleRate = 96000;
        [DataMember] public int InputCustomBitDepth = 24;
        [DataMember] public int MainResamplerQuality = 60;
        [DataMember] public int AuxResamplerQuality = 30;
        [DataMember] public bool MainForceInternalResamplerInShared = false;
        [DataMember] public bool AuxForceInternalResamplerInShared = false;
    }

    public sealed class StatusSnapshot
    {
        public bool InputExclusive;
        public bool Running;
        public string InputRole, InputFormat, InputDevice;
        public string InputRequested, InputAccepted, InputMix;
        public string MainDevice, AuxDevice, MainMode, AuxMode, MainSync, AuxSync, MainFormat, AuxFormat;

        public int MainBufferRequestedMs, AuxBufferRequestedMs;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs, AuxDefaultPeriodMs, AuxMinimumPeriodMs;
        public double MainAlignedMultiple, AuxAlignedMultiple;

        public bool MainNoSRC, AuxNoSRC, MainResampling, AuxResampling;
        public bool MainInternalResampler, AuxInternalResampler;
        public int MainInternalResamplerQuality, AuxInternalResamplerQuality;
        public bool MainMultiSRC, AuxMultiSRC;
    }

    static class Config
    {
        static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
        static readonly string FilePath = Path.Combine(Dir, "settings.json");

        public static string AppDataDir => Dir;
        public static string LogPath => Path.Combine(Dir, "MirrorAudio.log");

        public static AppSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new AppSettings();
                using (var fs = File.OpenRead(FilePath))
                    return (AppSettings)new DataContractJsonSerializer(typeof(AppSettings)).ReadObject(fs);
            }
            catch { return new AppSettings(); }
        }

        public static void Save(AppSettings s)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                using (var fs = File.Create(FilePath))
                    new DataContractJsonSerializer(typeof(AppSettings)).WriteObject(fs, s);
            }
            catch { }
        }
    }

    static class Log
    {
        static object _lock = new object();
        public static bool Enabled = false;
        public static void Info(string msg)
        {
            if (!Enabled) return;
            try
            {
                if (!Directory.Exists(Config.AppDataDir)) Directory.CreateDirectory(Config.AppDataDir);
                lock (_lock)
                {
                    File.AppendAllText(Config.LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}\r\n");
                }
            }
            catch { }
        }
        public static void Error(string msg, Exception ex = null)
        {
            Info("ERROR: " + msg + (ex != null ? " | " + ex : ""));
        }
    }

    static class AutoStartHelper
    {
        const string RUN_KEY = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string VALUE_NAME = "MirrorAudio";

        public static void Apply(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RUN_KEY, true) ?? Registry.CurrentUser.CreateSubKey(RUN_KEY))
                {
                    if (enable)
                    {
                        string exe = Application.ExecutablePath;
                        string val = $"\"{exe}\"";
                        key.SetValue(VALUE_NAME, val, RegistryValueKind.String);
                        Log.Info("AutoStart enabled: " + val);
                    }
                    else
                    {
                        if (key.GetValue(VALUE_NAME) != null)
                        {
                            key.DeleteValue(VALUE_NAME, false);
                            Log.Info("AutoStart disabled");
                        }
                    }
                }
            }
            catch (Exception ex) { Log.Error("Apply AutoStart failed", ex); }
        }
    }

    sealed class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture; BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux; WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;

        bool _running;
        bool _mainExclusive, _auxExclusive;
        bool _mainEventSyncUsed, _auxEventSyncUsed;

        bool _inExclusive = false;
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        string _inReqStr = "-", _inAccStr = "-", _inMixStr = "-";
        bool _mainNoSRC, _auxNoSRC, _mainResampling, _auxResampling;
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(4);

        public TrayApp()
        {
            Log.Enabled = _cfg.EnableLogging;
            AutoStartHelper.Apply(_cfg.AutoStart);
            Log.Info("App started. Logging=" + Log.Enabled + ", AutoStart=" + _cfg.AutoStart);

            try { _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var miSet = new ToolStripMenuItem("设置(&G)...", null, (s, e) => OnSettings());
            var miExit = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            StartOrRestart();
        }

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result; Config.Save(_cfg);
                    Log.Enabled = _cfg.EnableLogging;
                    Log.Info("Settings saved. Logging=" + Log.Enabled + ", AutoStart=" + _cfg.AutoStart + ", InputExclusive=" + _cfg.InputExclusive);
                    AutoStartHelper.Apply(_cfg.AutoStart);
                    StartOrRestart();
                }
            }
        }

        void StartOrRestart()
        {
            Log.Info("StartOrRestart()");
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            _inDev = FirstNonNull(FindById(_cfg.InputDeviceId, DataFlow.Capture),
                                  FindById(_cfg.InputDeviceId, DataFlow.Render),
                                  _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia));
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux = FindById(_cfg.AuxDeviceId, DataFlow.Render);
            _inDevName = _inDev != null ? _inDev.FriendlyName : "-";

            Log.Info("InputDev=" + _inDevName + ", MainOut=" + SafeName(_cfg.MainDeviceId, DataFlow.Render) + ", AuxOut=" + SafeName(_cfg.AuxDeviceId, DataFlow.Render));

            if (_outMain == null || _outAux == null)
            {
                MessageBox.Show("请在“设置”选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Log.Info("Missing main/aux device, abort.");
                return;
            }

            WaveFormat inFmt;
            WaveFormat inputRequested = null, inputAccepted = null, inputMix = null;

            if (_inDev.DataFlow == DataFlow.Capture)
            {
                _inRoleStr = "录音";
                try { inputMix = _inDev.AudioClient.MixFormat; } catch { }
                var req = new InputFormatRequest
                {
                    Strategy = _cfg.InputFormatStrategy,
                    CustomSampleRate = _cfg.InputCustomSampleRate,
                    CustomBitDepth = _cfg.InputCustomBitDepth,
                    Channels = 2
                };

                IWaveIn cap = null;
                _inExclusive = false;

                try
                {
                    var desired = InputFormatHelper.BuildWaveFormat(req.Strategy, req.CustomSampleRate, req.CustomBitDepth, 2);
                    inputRequested = desired;
                    if (_cfg.InputExclusive && desired != null)
                    {
                        cap = TryCreateExclusiveCapture(_inDev, desired, out inputAccepted);
                    }
                    if (cap == null)
                    {
                        var sharedCap = new WasapiCapture(_inDev);
                        if (desired != null) sharedCap.WaveFormat = desired;
                        cap = sharedCap;
                        inputAccepted = sharedCap.WaveFormat;
                    }
                    _capture = cap; inFmt = inputAccepted ?? desired ?? inputMix;
                }
                catch (Exception ex)
                {
                    Log.Error("Create capture failed", ex);
                    MessageBox.Show("创建输入设备失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisposeAll(); return;
                }

                _inReqStr = InputFormatHelper.Fmt(inputRequested);
                _inAccStr = InputFormatHelper.Fmt(inFmt);
                _inMixStr = InputFormatHelper.Fmt(inputMix);
            }
            else
            {
                _inRoleStr = "环回";
                var cap = new WasapiLoopbackCapture(_inDev);
                string negoLog;
                var req = new InputFormatRequest
                {
                    Strategy = _cfg.InputFormatStrategy,
                    CustomSampleRate = _cfg.InputCustomSampleRate,
                    CustomBitDepth = _cfg.InputCustomBitDepth,
                    Channels = 2
                };
                var wf = InputFormatHelper.NegotiateLoopbackFormat(_inDev, req, out negoLog, out inputMix, out inputAccepted, out inputRequested);
                if (wf != null) cap.WaveFormat = wf;
                _capture = cap; inFmt = cap.WaveFormat;
                _inExclusive = false;
                Log.Info("Loopback: " + negoLog?.Replace("\r", " ").Replace("\n", " | "));
                _inReqStr = InputFormatHelper.Fmt(inputRequested);
                _inAccStr = InputFormatHelper.Fmt(inputAccepted ?? inFmt);
                _inMixStr = InputFormatHelper.Fmt(inputMix);
            }

            _inFmtStr = Fmt(inFmt);

            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 4, 80)) };
            _bufAux  = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs * 4, 120)) };

            ContinueStart(inFmt);
        }

        void ContinueStart(WaveFormat inFmt)
        {
            GetPeriods(_outMain, out _defMainMs, out _minMainMs);
            GetPeriods(_outAux,  out _defAuxMs,  out _minAuxMs);
            Log.Info($"Periods Main(def={_defMainMs:0.###} min={_minMainMs:0.###}) Aux(def={_defAuxMs:0.###} min={_minAuxMs:0.###})");

            _srcMain = _bufMain; _resMain = null; _mainExclusive = false; _mainEventSyncUsed = false; _mainBufEffectiveMs = _cfg.MainBufMs; _mainFmtStr = "-";
            _mainNoSRC = false; _mainResampling = false;
            var desiredMain = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);

            bool isLoopMain = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outMain.ID;
            bool wantExMain = (_cfg.MainShare == ShareModeOption.Exclusive || _cfg.MainShare == ShareModeOption.Auto) && !isLoopMain;
            if (isLoopMain && (_cfg.MainShare != ShareModeOption.Shared))
            {
                MessageBox.Show("输入为主设备环回，独占冲突，主通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Log.Info("Main exclusive disabled due to loopback input conflict.");
            }

            WaveFormat mainTargetFmt = null;
            if (wantExMain && SupportsExclusive(_outMain, desiredMain))
            {
                bool needChange = (inFmt.SampleRate != desiredMain.SampleRate) || (inFmt.Channels != desiredMain.Channels);
                if (needChange) _srcMain = _resMain = new MediaFoundationResampler(_bufMain, desiredMain) { ResamplerQuality = _cfg.MainResamplerQuality };
                int ms = BufAligned(_cfg.MainBufMs, true, _defMainMs, _minMainMs, _cfg.MainBufMode);
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                if (_mainOut != null)
                {
                    _mainExclusive = true; _mainBufEffectiveMs = ms; _mainFmtStr = Fmt(desiredMain); mainTargetFmt = desiredMain;
                    _mainResampling = needChange; _mainNoSRC = !needChange;
                    Log.Info($"Main Out EXCLUSIVE {Fmt(desiredMain)} buf={ms}ms resample={needChange}");
                }
            }
            if (_mainOut == null)
            {
                int ms = BufAligned(_cfg.MainBufMs, false, _defMainMs, 0, _cfg.MainBufMode);
                WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }

                if (_cfg.MainForceInternalResamplerInShared && mix != null)
                {
                    bool needChange = (inFmt.SampleRate != mix.SampleRate) || (inFmt.Channels != mix.Channels) || (inFmt.BitsPerSample != mix.BitsPerSample);
                    if (needChange)
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, mix) { ResamplerQuality = _cfg.MainResamplerQuality };
                        _srcMain = _resMain;
                        _mainResampling = true; _mainNoSRC = false;
                        Log.Info($"Main Shared internal resampler -> {Fmt(mix)} quality={_cfg.MainResamplerQuality}");
                    }
                }

                _mainOut = CreateOut(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                if (_mainOut == null)
                {
                    MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Log.Error("Main Out init failed");
                    DisposeAll(); return;
                }
                _mainBufEffectiveMs = ms;
                try { mainTargetFmt = _outMain.AudioClient.MixFormat; _mainFmtStr = Fmt(mainTargetFmt); } catch { _mainFmtStr = "系统混音"; }
                if (_resMain == null)
                {
                    _mainResampling = (inFmt.SampleRate != (mainTargetFmt != null ? mainTargetFmt.SampleRate : inFmt.SampleRate) ||
                                       inFmt.Channels    != (mainTargetFmt != null ? mainTargetFmt.Channels    : inFmt.Channels));
                    _mainNoSRC = !_mainResampling;
                }
                Log.Info($"Main Out SHARED {(_mainFmtStr)} buf={ms}ms resample={_mainResampling}");
            }

            _srcAux = _bufAux; _resAux = null; _auxExclusive = false; _auxEventSyncUsed = false; _auxBufEffectiveMs = _cfg.AuxBufMs; _auxFmtStr = "-";
            _auxNoSRC = false; _auxResampling = false;
            var desiredAux = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);

            bool isLoopAux = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outAux.ID;
            bool wantExAux = (_cfg.AuxShare == ShareModeOption.Exclusive || _cfg.AuxShare == ShareModeOption.Auto) && !isLoopAux;
            if (isLoopAux && (_cfg.AuxShare != ShareModeOption.Shared))
            {
                MessageBox.Show("输入为副设备环回，独占冲突，副通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Log.Info("Aux exclusive disabled due to loopback input conflict.");
            }

            WaveFormat auxTargetFmt = null;
            if (wantExAux && SupportsExclusive(_outAux, desiredAux))
            {
                bool needChange = (inFmt.SampleRate != desiredAux.SampleRate) || (inFmt.Channels != desiredAux.Channels);
                if (needChange) _srcAux = _resAux = new MediaFoundationResampler(_bufAux, desiredAux) { ResamplerQuality = _cfg.AuxResamplerQuality };
                int ms = BufAligned(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs, _cfg.AuxBufMode);
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, ms, _srcAux, out _auxEventSyncUsed);
                if (_auxOut != null)
                {
                    _auxExclusive = true; _auxBufEffectiveMs = ms; _auxFmtStr = Fmt(desiredAux); auxTargetFmt = desiredAux;
                    _auxResampling = needChange; _auxNoSRC = !needChange;
                    Log.Info($"Aux Out EXCLUSIVE {Fmt(desiredAux)} buf={ms}ms resample={needChange}");
                }
            }
            if (_auxOut == null)
            {
                int ms = BufAligned(_cfg.AuxBufMs, false, _defAuxMs, 0, _cfg.AuxBufMode);
                WaveFormat mix = null; try { mix = _outAux.AudioClient.MixFormat; } catch { }

                if (_cfg.AuxForceInternalResamplerInShared && mix != null)
                {
                    bool needChange = (inFmt.SampleRate != mix.SampleRate) || (inFmt.Channels != mix.Channels) || (inFmt.BitsPerSample != mix.BitsPerSample);
                    if (needChange)
                    {
                        _resAux = new MediaFoundationResampler(_bufAux, mix) { ResamplerQuality = _cfg.AuxResamplerQuality };
                        _srcAux = _resAux;
                        _auxResampling = true; _auxNoSRC = false;
                        Log.Info($"Aux Shared internal resampler -> {Fmt(mix)} quality={_cfg.AuxResamplerQuality}");
                    }
                }

                _auxOut = CreateOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, ms, _srcAux, out _auxEventSyncUsed);
                if (_auxOut == null)
                {
                    MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Log.Error("Aux Out init failed");
                    DisposeAll(); return;
                }
                _auxBufEffectiveMs = ms;
                try { auxTargetFmt = _outAux.AudioClient.MixFormat; _auxFmtStr = Fmt(auxTargetFmt); } catch { _auxFmtStr = "系统混音"; }
                if (_resAux == null)
                {
                    _auxResampling = (inFmt.SampleRate != (auxTargetFmt != null ? auxTargetFmt.SampleRate : inFmt.SampleRate) ||
                                      inFmt.Channels    != (auxTargetFmt != null ? auxTargetFmt.Channels    : inFmt.Channels));
                    _auxNoSRC = !_auxResampling;
                }
                Log.Info($"Aux Out SHARED {(_auxFmtStr)} buf={ms}ms resample={_auxResampling}");
            }

            _capture.DataAvailable += OnIn; _capture.RecordingStopped += OnStopRec;
            try
            {
                _capture.StartRecording();
                _mainOut.Play(); _auxOut.Play(); _running = true;
                Log.Info("Started: capture+main+aux running.");
            }
            catch (Exception ex)
            {
                Log.Error("Start failed", ex);
                MessageBox.Show("启动失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void OnIn(object s, WaveInEventArgs e)
        {
            if (_bufMain != null) _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded);
            if (_bufAux  != null) _bufAux .AddSamples(e.Buffer, 0, e.BytesRecorded);
        }
        void OnStopRec(object s, StoppedEventArgs e)
        {
            try { if (_bufMain != null) _bufMain.ClearBuffer(); if (_bufAux != null) _bufAux.ClearBuffer(); } catch { }
            Log.Info("Capture stopped.");
        }

        public void Stop()
        {
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut .Stop(); } catch { }
            Thread.Sleep(20);
            DisposeAll();
            _running = false;
            Log.Info("Stopped & disposed.");
        }

        void DisposeAll()
        {
            try
            {
                if (_capture != null) { _capture.DataAvailable -= OnIn; _capture.RecordingStopped -= OnStopRec; _capture.Dispose(); }
            } catch { } _capture = null;

            try { _mainOut?.Dispose(); } catch { } _mainOut = null;
            try { _auxOut ?.Dispose(); } catch { } _auxOut  = null;
            try { _resMain?.Dispose(); } catch { } _resMain = null;
            try { _resAux ?.Dispose(); } catch { } _resAux  = null;
            _bufMain = null; _bufAux = null;
        }

        public StatusSnapshot GetStatusSnapshot()
        {
            bool mainInternal = _resMain != null;
            bool auxInternal  = _resAux  != null;
            bool mainMulti = false, auxMulti = false;
            int mainQ = 0, auxQ = 0;

            try
            {
                if (mainInternal) mainQ = _resMain.ResamplerQuality;
                if (auxInternal)  auxQ  = _resAux .ResamplerQuality;

                if (!_mainExclusive && mainInternal)
                {
                    WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }
                    if (mix != null)
                    {
                        var f = _resMain.WaveFormat;
                        if (f.SampleRate != mix.SampleRate || f.Channels != mix.Channels || f.BitsPerSample != mix.BitsPerSample) mainMulti = true;
                    }
                }
                if (!_auxExclusive && auxInternal)
                {
                    WaveFormat mix = null; try { mix = _outAux.AudioClient.MixFormat; } catch { }
                    if (mix != null)
                    {
                        var f = _resAux.WaveFormat;
                        if (f.SampleRate != mix.SampleRate || f.Channels != mix.Channels || f.BitsPerSample != mix.BitsPerSample) auxMulti = true;
                    }
                }
            }
            catch { }

            double mainStep = (_cfg.MainBufMode == BufferAlignMode.MinAlign ? (_minMainMs > 0 ? _minMainMs : _defMainMs) : _defMainMs);
            double auxStep  = (_cfg.AuxBufMode  == BufferAlignMode.MinAlign ? (_minAuxMs  > 0 ? _minAuxMs  : _defAuxMs ) : _defAuxMs );
            double mainMul  = (_mainBufEffectiveMs > 0 && mainStep > 0) ? _mainBufEffectiveMs / mainStep : 0;
            double auxMul   = (_auxBufEffectiveMs  > 0 && auxStep  > 0) ? _auxBufEffectiveMs  / auxStep  : 0;

            return new StatusSnapshot
            {
                InputExclusive = _inExclusive, Running = _running,
                InputRole = _inRoleStr, InputFormat = _inFmtStr, InputDevice = _inDevName,
                InputRequested = _inReqStr, InputAccepted = _inAccStr, InputMix = _inMixStr,

                MainDevice = _outMain != null ? _outMain.FriendlyName : SafeName(_cfg.MainDeviceId, DataFlow.Render),
                AuxDevice  = _outAux  != null ? _outAux .FriendlyName : SafeName(_cfg.AuxDeviceId,  DataFlow.Render),

                MainMode = _mainOut != null ? (_mainExclusive ? "独占" : "共享") : "-",
                AuxMode  = _auxOut  != null ? (_auxExclusive  ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEventSyncUsed ? "事件" : "轮询") : "-",
                AuxSync  = _auxOut  != null ? (_auxEventSyncUsed  ? "事件" : "轮询") : "-",

                MainFormat = _mainOut != null ? _mainFmtStr : "-",
                AuxFormat  = _auxOut  != null ? _auxFmtStr  : "-",

                MainBufferRequestedMs = _cfg.MainBufMs, AuxBufferRequestedMs = _cfg.AuxBufMs,
                MainBufferMs = _mainOut != null ? _mainBufEffectiveMs : 0,
                AuxBufferMs  = _auxOut  != null ? _auxBufEffectiveMs  : 0,

                MainDefaultPeriodMs = _defMainMs, MainMinimumPeriodMs = _minMainMs,
                AuxDefaultPeriodMs  = _defAuxMs,  AuxMinimumPeriodMs  = _minAuxMs,

                MainAlignedMultiple = mainMul, AuxAlignedMultiple = auxMul,

                MainNoSRC = _mainNoSRC, AuxNoSRC = _auxNoSRC,
                MainResampling = _mainResampling, AuxResampling = _auxResampling,

                MainInternalResampler = mainInternal, AuxInternalResampler = auxInternal,
                MainInternalResamplerQuality = mainQ, AuxInternalResamplerQuality = auxQ,
                MainMultiSRC = mainMulti, AuxMultiSRC = auxMulti
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
        static T FirstNonNull<T>(params T[] arr) where T : class { foreach (var a in arr) if (a != null) return a; return null; }
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
            double stepMin = (minMs > 0 ? minMs : (defMs > 0 ? defMs : 10.0));
            double stepDef = (defMs > 0 ? defMs : stepMin);
            int ms;
            if (exclusive)
            {
                if (mode == BufferAlignMode.MinAlign) ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepMin) * stepMin);
                else                                   ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepDef) * stepDef);
                double floor = (mode == BufferAlignMode.MinAlign ? stepMin : stepDef) * 3.0;
                if (ms < floor)
                {
                    double step = (mode == BufferAlignMode.MinAlign ? stepMin : stepDef);
                    ms = (int)Math.Ceiling(Math.Ceiling(floor / step) * step);
                }
                return ms;
            }
            else
            {
                if (mode == BufferAlignMode.MinAlign) ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepMin) * stepMin);
                else                                   ms = (int)Math.Ceiling(Math.Ceiling(wantMs / stepDef) * stepDef);
                double floor = stepDef * 2.0;
                if (ms < floor)
                {
                    double step = (mode == BufferAlignMode.MinAlign ? stepMin : stepDef);
                    ms = (int)Math.Ceiling(Math.Ceiling(floor / step) * step);
                }
                return ms;
            }
        }

        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption syncMode, int bufMs, IWaveProvider src, out bool eventSync)
        {
            eventSync = false;
            try
            {
                var useEvent = (syncMode == SyncModeOption.Event || (syncMode == SyncModeOption.Auto));
                var wo = new WasapiOut(dev, mode, useEvent ? true : false, bufMs);
                eventSync = useEvent;
                wo.Init(src);
                return wo;
            }
            catch (Exception ex) { Log.Error("CreateOut failed", ex); return null; }
        }

        IWaveIn TryCreateExclusiveCapture(MMDevice dev, WaveFormat req, out WaveFormat accepted)
        {
            accepted = null;
            try
            {
                var cap = new WasapiCapture(dev, true, 20, AudioClientShareMode.Exclusive);
                if (req != null) ((WasapiCapture)cap).WaveFormat = req;
                accepted = ((WasapiCapture)cap).WaveFormat;
                _inExclusive = true;
                return cap;
            }
            catch
            {
                return null;
            }
        }
    
        public void Dispose()
        {
            try { Stop(); } catch { }
            try { if (_tray != null) { _tray.Visible = false; _tray.Dispose(); } } catch { }
            try { _menu?.Dispose(); } catch { }
            try { _mm?.Dispose(); } catch { }
        }
}

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
                case InputFormatStrategy.SystemMix:       return null;
                case InputFormatStrategy.Specify24_48000: return CreatePcm24(48000, channels);
                case InputFormatStrategy.Specify24_96000: return CreatePcm24(96000, channels);
                case InputFormatStrategy.Specify24_192000:return CreatePcm24(192000, channels);
                case InputFormatStrategy.Specify32f_48000:return WaveFormat.CreateIeeeFloatWaveFormat(48000, channels);
                case InputFormatStrategy.Specify32f_96000:return WaveFormat.CreateIeeeFloatWaveFormat(96000, channels);
                case InputFormatStrategy.Specify32f_192000:return WaveFormat.CreateIeeeFloatWaveFormat(192000, channels);
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

        public static string Fmt(WaveFormat wf) { return wf == null ? "-" : (wf.SampleRate + "Hz/" + wf.BitsPerSample + "bit/" + wf.Channels + "ch"); }

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
            try { ok = device.AudioClient.IsFormatSupported(AudioClientShareMode.Shared, desired, out closest); } catch { ok = false; }
            sb.AppendLine("Request: " + Fmt(desired) + " -> Supported: " + (ok ? "Yes" : "No"));

            if (!ok && closest != null)
            {
                acceptedFormat = closest;
                sb.AppendLine("Closest: " + Fmt(closest));
                log = sb.ToString();
                return closest;
            }
            if (ok)
            {
                acceptedFormat = desired;
                log = sb.ToString();
                return desired;
            }
            log = sb.ToString();
            return null;
        }
    }
}
