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

        // 内部重采样质量（60/50/40/30），共享下也程序内重采样
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

    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        void ApplySideEffects() { StartupHelper.SetAutoStart(_cfg.AutoStart); Logger.Init(_cfg.EnableLogging); }
        bool _inExclusive = false;
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
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        string _inReqStr = "-", _inAccStr = "-", _inMixStr = "-";
        bool _mainNoSRC, _auxNoSRC, _mainResampling, _auxResampling;
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(4);

        public TrayApp()
        {
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }
            try { _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var miSet = new ToolStripMenuItem("设置(&G)...", null, (s, e) => OnSettings());
            var miExit = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu; ApplySideEffects(); Logger.Log("TrayApp constructed.");

            StartOrRestart();
        }

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result; Config.Save(_cfg);
                    StartOrRestart();
                }
            }
        }

        // IMMNotificationClient
        public void OnDeviceStateChanged(string id, DeviceState st) { }
        public void OnDeviceAdded(string id) { }
        public void OnDeviceRemoved(string id) { }
        public void OnDefaultDeviceChanged(DataFlow f, Role r, string id) { }
        public void OnPropertyValueChanged(string id, PropertyKey key) { }

        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            _inDev = FirstNonNull(FindById(_cfg.InputDeviceId, DataFlow.Capture),
                                  FindById(_cfg.InputDeviceId, DataFlow.Render),
                                  _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia));
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux = FindById(_cfg.AuxDeviceId, DataFlow.Render);
            _inDevName = _inDev != null ? _inDev.FriendlyName : "-";

            if (_outMain == null || _outAux == null)
            {
                MessageBox.Show("请在“设置”选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                var acc = InputFormatHelper.BuildWaveFormat(req.Strategy, req.CustomSampleRate, req.CustomBitDepth, 2);
                _inExclusive = _cfg.InputExclusive;
                WaveFormat inputAccepted2 = null;
                if (_inExclusive)
                {
                    cap = TryCreateExclusiveCapture(_inDev, acc, out inputAccepted);
                    if (cap == null) _inExclusive = false;
                }
                if (cap == null)
                {
                    var shared = new WasapiCapture(_inDev);
                    if (acc != null) shared.WaveFormat = acc;
                    inputAccepted2 = shared.WaveFormat;
                    cap = shared;
                }
                _capture = cap; inFmt = cap.WaveFormat;

                _inReqStr = InputFormatHelper.Fmt(acc);
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
                _inReqStr = InputFormatHelper.Fmt(inputRequested);
                _inAccStr = InputFormatHelper.Fmt(inputAccepted ?? inputAccepted2 ?? inFmt);
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

            // ========== 主通道 ==========
            _srcMain = _bufMain; _resMain = null; _mainExclusive = false; _mainEventSyncUsed = false; _mainBufEffectiveMs = _cfg.MainBufMs; _mainFmtStr = "-";
            _mainNoSRC = false; _mainResampling = false;
            var desiredMain = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);

            bool isLoopMain = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outMain.ID;
            bool wantExMain = (_cfg.MainShare == ShareModeOption.Exclusive || _cfg.MainShare == ShareModeOption.Auto) && !isLoopMain;
            if (isLoopMain && (_cfg.MainShare != ShareModeOption.Shared))
                MessageBox.Show("输入为主设备环回，独占冲突，主通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
                }
            }
            if (_mainOut == null)
            {
                int ms = BufAligned(_cfg.MainBufMs, false, _defMainMs, 0, _cfg.MainBufMode);
                WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }

                // 共享模式下也程序内重采样（若开启且输入 != MixFormat）
                if (_cfg.MainForceInternalResamplerInShared && mix != null)
                {
                    bool needChange = (inFmt.SampleRate != mix.SampleRate) || (inFmt.Channels != mix.Channels) || (inFmt.BitsPerSample != mix.BitsPerSample);
                    if (needChange)
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, mix) { ResamplerQuality = _cfg.MainResamplerQuality };
                        _srcMain = _resMain;
                        _mainResampling = true; _mainNoSRC = false;
                    }
                }

                _mainOut = CreateOut(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                if (_mainOut == null)
                {
                    MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            }

            // ========== 副通道 ==========
            _srcAux = _bufAux; _resAux = null; _auxExclusive = false; _auxEventSyncUsed = false; _auxBufEffectiveMs = _cfg.AuxBufMs; _auxFmtStr = "-";
            _auxNoSRC = false; _auxResampling = false;
            var desiredAux = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);

            bool isLoopAux = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outAux.ID;
            bool wantExAux = (_cfg.AuxShare == ShareModeOption.Exclusive || _cfg.AuxShare == ShareModeOption.Auto) && !isLoopAux;
            if (isLoopAux && (_cfg.AuxShare != ShareModeOption.Shared))
                MessageBox.Show("输入为副设备环回，独占冲突，副通道改走共享。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
                    }
                }

                _auxOut = CreateOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, ms, _srcAux, out _auxEventSyncUsed);
                if (_auxOut == null)
                {
                    MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            }

            _capture.DataAvailable += OnIn; _capture.RecordingStopped += OnStopRec;
            try
            {
                _capture.StartRecording();
                _mainOut.Play(); _auxOut.Play(); _running = true;
            }
            catch (Exception)
            {
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
        }

        public void Stop()
        {
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut .Stop(); } catch { }
            Thread.Sleep(20);
            DisposeAll();
            _running = false;
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

        public void Dispose()
        {
            try { Stop(); } catch { }
            try { if (_mm != null) { _mm.UnregisterEndpointNotificationCallback(this); _mm.Dispose(); } } catch { }
            try { if (_tray != null) { _tray.Visible = false; _tray.Dispose(); } } catch { }
            try { _menu?.Dispose(); } catch { }
            GC.SuppressFinalize(this);
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

            // 与 BufAligned 同步的倍数计算（依据所选对齐模式）
            double mainStep = (_cfg.MainBufMode == BufferAlignMode.MinAlign ? (_minMainMs > 0 ? _minMainMs : _defMainMs) : _defMainMs);
            double auxStep  = (_cfg.AuxBufMode  == BufferAlignMode.MinAlign ? (_minAuxMs  > 0 ? _minAuxMs  : _defAuxMs ) : _defAuxMs );
            double mainMul  = (_mainBufEffectiveMs > 0 && mainStep > 0) ? _mainBufEffectiveMs / mainStep : 0;
            double auxMul   = (_auxBufEffectiveMs  > 0 && auxStep  > 0) ? _auxBufEffectiveMs  / auxStep  : 0;

            return new StatusSnapshot
            {
                InputExclusive = _inExclusive,
                Running = _running,
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
                // 独占：至少 3× 步长
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
                // 共享：至少 2× 默认周期
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
            catch { return null; }
        }
    
        // Try to create exclusive capture; return null to fallback
        IWaveIn TryCreateExclusiveCapture(MMDevice dev, WaveFormat req, out WaveFormat accepted)
        {
            accepted = null;
            try
            {
                var cap = new ExclusiveWasapiCapture(dev, req);
                accepted = cap.WaveFormat;
                return cap;
            }
            catch
            {
                return null;
            }
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

    // Minimal exclusive-mode capture using NAudio CoreAudio
    class ExclusiveWasapiCapture : IWaveIn
    {
        public event EventHandler<WaveInEventArgs> DataAvailable;
        public event EventHandler<StoppedEventArgs> RecordingStopped;

        readonly MMDevice _device;
        readonly WaveFormat _request;
        WaveFormat _format;
        AudioClient _client;
        AudioCaptureClient _capture;
        IntPtr _eventHandle = IntPtr.Zero;
        Thread _thread;
        volatile bool _stop;

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        static extern bool SetEvent(IntPtr hEvent);
        const uint INFINITE = 0xFFFFFFFF;

        public ExclusiveWasapiCapture(MMDevice device, WaveFormat requestFormat = null)
        {
            _device = device ?? throw new ArgumentNullException(nameof(device));
            _request = requestFormat ?? new WaveFormat(192000, 16, 2);
            Initialize();
        }

        void Initialize()
        {
            _client = _device.AudioClient;
            long defPer, minPer;
            _client.DefaultDevicePeriod; var minPer = _client.MinimumDevicePeriod; var defPer = _client.DefaultDevicePeriod;

            var fmt = _request;
            if (!_client.IsFormatSupported(AudioClientShareMode.Exclusive, fmt))
            {
                var trials = new List<WaveFormat>{
                    new WaveFormat(192000, 24, 2),
                    WaveFormat.CreateIeeeFloatWaveFormat(192000, 2),
                    new WaveFormat(96000, 16, 2),
                    new WaveFormat(48000, 16, 2)
                };
                foreach (var f in trials) { if (_client.IsFormatSupported(AudioClientShareMode.Exclusive, f)) { fmt = f; break; } }
            }
            _format = fmt;
            _client.Initialize(AudioClientShareMode.Exclusive, AudioClientStreamFlags.EventCallback, minPer, minPer, _format, Guid.Empty);
            var _bsz = _client.BufferSize;
            _capture = _client.AudioCaptureClient;
            _eventHandle = CreateEvent(IntPtr.Zero, false, false, null);
            _client.SetEventHandle(_eventHandle);
        }

        public WaveFormat WaveFormat { get => _format; set => throw new NotSupportedException(); }

        public void StartRecording()
        {
            if (_thread != null) return;
            _stop = false;
            _client.Start();
            _thread = new Thread(CaptureThread) { IsBackground = true, Priority = ThreadPriority.Highest };
            _thread.Start();
        }

        void CaptureThread()
        {
            try
            {
                while (!_stop)
                {
                    WaitForSingleObject(_eventHandle, INFINITE);
                    if (_stop) break;
                    int next;
                    next = _capture.GetNextPacketSize();
                    while (next > 0)
                    {
                        IntPtr ptr;
                        int frames;
                        AudioClientBufferFlags flags;
                        long devpos, qpc;
                        ptr = _capture.GetBuffer(out frames, out flags, out devpos, out qpc);
                        int bytes = frames * _format.BlockAlign;
                        if (bytes > 0)
                        {
                            byte[] buf = new byte[bytes];
                            System.Runtime.InteropServices.Marshal.Copy(ptr, buf, 0, bytes);
                            DataAvailable?.Invoke(this, new WaveInEventArgs(buf, bytes));
                        }
                        _capture.ReleaseBuffer(frames);
                        next = _capture.GetNextPacketSize();
                    }
                }
            }
            catch (Exception ex) { RecordingStopped?.Invoke(this, new StoppedEventArgs(ex)); }
            finally { try { _client.Stop(); } catch { } }
        }

        public void StopRecording()
        {
            _stop = true;
            try { SetEvent(_eventHandle); } catch { }
            try { _thread?.Join(500); } catch { }
            _thread = null;
        }

        public void Dispose()
        {
            StopRecording();
        }
    }

    static class Logger
    {
        static bool _enabled; 
        static readonly object _lock = new object();
        static string Dir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio", "logs");
        static string FilePath => Path.Combine(Dir, DateTime.Now.ToString("yyyyMMdd") + ".log");
        public static void Init(bool enabled)
        {
            _enabled = enabled;
            if (_enabled && !Directory.Exists(Dir)) { try { Directory.CreateDirectory(Dir); } catch {} }
        }
        public static void Log(string msg)
        {
            if (!_enabled) return;
            try { lock(_lock) System.IO.File.AppendAllText(FilePath, DateTime.Now.ToString("HH:mm:ss.fff ") + msg + Environment.NewLine); } catch {}
        }
    }
    static class StartupHelper
    {
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string AppName = "MirrorAudio";
        public static void SetAutoStart(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RunKey, true) ?? Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (enable)
                    {
                        var exe = Application.ExecutablePath;
                        key.SetValue(AppName, $"\"{exe}\"");
                    }
                    else if (key.GetValue(AppName) != null)
                    {
                        key.DeleteValue(AppName, false);
                    }
                }
            }
            catch {}
        }
    }
}
