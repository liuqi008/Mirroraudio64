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
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace MirrorAudio
{
    // ---------- Logger ----------
    static class Logger
    {
        public static bool Enabled;
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        static readonly string CrashPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
        public static void Info(string s) { if (!Enabled) return; try { File.AppendAllText(LogPath, "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + s + "\r\n"); } catch { } }
        public static void Crash(string where, Exception ex) { if (ex == null) return; try { File.AppendAllText(CrashPath, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + where + ": " + ex + "\r\n"); } catch { } }
    }

    // ---------- Program ----------
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.ThreadException += (s, e) => Logger.Crash("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Crash("NonUI", e.ExceptionObject as Exception);
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

    // ---------- Enums & Settings ----------
    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }
    [DataContract] public enum BufferAlignMode { [EnumMember] DefaultAlign, [EnumMember] MinAlign }

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
    public sealed class AppSettings
    {
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption MainSync = SyncModeOption.Auto, AuxSync = SyncModeOption.Auto;
        [DataMember] public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        [DataMember] public int AuxRate = 48000, AuxBits = 16, AuxBufMs = 150;
        [DataMember] public bool AutoStart = false, EnableLogging = false;
        [DataMember] public BufferAlignMode MainBufMode = BufferAlignMode.DefaultAlign;
        [DataMember] public BufferAlignMode AuxBufMode = BufferAlignMode.DefaultAlign;

        // Input loopback strategy
        [DataMember] public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        [DataMember] public int InputCustomSampleRate = 96000;
        [DataMember] public int InputCustomBitDepth = 24;

        // Resampler
        [DataMember] public int MainResamplerQuality = 50;  // 60/50/40/30
        [DataMember] public int AuxResamplerQuality = 30;
        [DataMember] public bool MainForceInternalResamplerInShared = false;
        [DataMember] public bool AuxForceInternalResamplerInShared = false;
    }

    // ---------- Status Snapshot ----------
    public sealed class StatusSnapshot
    {
        public bool Running;
        public string InputRole, InputFormat, InputDevice;
        public string InputRequested, InputAccepted, InputMix;
        public string MainDevice, AuxDevice, MainMode, AuxMode, MainSync, AuxSync, MainFormat, AuxFormat;
        public int MainBufferRequestedMs, AuxBufferRequestedMs;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs, AuxDefaultPeriodMs, AuxMinimumPeriodMs;
        public double MainAlignedMultiple, AuxAlignedMultiple;
        public double MainBufferMultiple, AuxBufferMultiple;
        public bool MainNoSRC, AuxNoSRC, MainResampling, AuxResampling;

        // New fields
        public bool MainInternalResampler, AuxInternalResampler;
        public bool MainMultiSRC, AuxMultiSRC;
        public int MainInternalResamplerQuality; // -1 if not active
        public int AuxInternalResamplerQuality;  // -1 if not active
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

    // ---------- Tray App ----------
    sealed class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture;
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux;
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;

        bool _running;
        bool _mainExclusive, _auxExclusive;
        bool _mainEvent, _auxEvent;
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        string _inReqStr = "-", _inAccStr = "-", _inMixStr = "-";
        bool _mainNoSRC, _auxNoSRC, _mainResampling, _auxResampling;
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        public TrayApp()
        {
            Logger.Enabled = _cfg.EnableLogging;
            SetupTray();
            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        void SetupTray()
        {
            _tray.Visible = true;
            _tray.Text = "MirrorAudio";
            try
            {
                _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var miSet = new ToolStripMenuItem("设置(&G)...", null, (s, e) => OnSettings());
            var miExit = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;
        }

        void OnSettings()
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
            }
            catch { }
        }

        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            _inDev = FindById(_cfg.InputDeviceId, DataFlow.Capture);
            if (_inDev == null) _inDev = FindById(_cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux = FindById(_cfg.AuxDeviceId, DataFlow.Render);
            if (_outMain == null || _outAux == null)
            {
                MessageBox.Show("请在“设置”选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _inDevName = _inDev != null ? _inDev.FriendlyName : "-";

            // ----- Input -----
            WaveFormat inFmt;
            WaveFormat inputRequested = null, inputAccepted = null, inputMix = null;
            try { inputMix = _inDev.AudioClient.MixFormat; } catch { }

            if (_inDev.DataFlow == DataFlow.Capture)
            {
                _inRoleStr = "录音";
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                var req = InputFormatHelper.BuildWaveFormat(_cfg.InputFormatStrategy, _cfg.InputCustomSampleRate, _cfg.InputCustomBitDepth, 2);
                if (req != null) cap.WaveFormat = req;
                inFmt = cap.WaveFormat;
                inputRequested = req;
                inputAccepted = inFmt;
                _capture = cap;
            }
            else
            {
                _inRoleStr = "环回";
                var cap = new WasapiLoopbackCapture(_inDev);
                var req = InputFormatHelper.BuildWaveFormat(_cfg.InputFormatStrategy, _cfg.InputCustomSampleRate, _cfg.InputCustomBitDepth, 2);
                if (req != null) cap.WaveFormat = req;
                inFmt = cap.WaveFormat;
                inputRequested = req ?? inputMix;
                inputAccepted = inFmt;
                _capture = cap;
            }

            _inFmtStr = Fmt(inFmt);
            _inReqStr = Fmt(inputRequested);
            _inAccStr = Fmt(inputAccepted);
            _inMixStr = Fmt(inputMix);

            // Buffers
            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 3.0, 60)) };
            _bufAux = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs * 4.0, 120)) };

            // Periods (best-effort; some drivers may not expose)
            GetPeriods(_outMain, out _defMainMs, out _minMainMs);
            GetPeriods(_outAux, out _defAuxMs, out _minAuxMs);

            // ----- Main Output -----
            _srcMain = _bufMain; _resMain = null; _mainExclusive = false; _mainEvent = false; _mainBufEffectiveMs = _cfg.MainBufMs; _mainFmtStr = "-";
            _mainNoSRC = false; _mainResampling = false;
            var desiredMain = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
            bool isLoopMain = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outMain.ID;
            bool wantExMain = (_cfg.MainShare == ShareModeOption.Exclusive || _cfg.MainShare == ShareModeOption.Auto) && !isLoopMain;

            WaveFormat mainTargetFmt = null;

            if (wantExMain && SupportsExclusive(_outMain, desiredMain))
            {
                bool needRateChange = (inFmt.SampleRate != desiredMain.SampleRate) || (inFmt.Channels != desiredMain.Channels);
                if (needRateChange) _srcMain = _resMain = new MediaFoundationResampler(_bufMain, desiredMain) { ResamplerQuality = _cfg.MainResamplerQuality };
                int ms = BufAligned(_cfg.MainBufMs, true, _defMainMs, _minMainMs, _cfg.MainBufMode);
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _srcMain, out _mainEvent);
                if (_mainOut != null)
                {
                    _mainExclusive = true;
                    _mainBufEffectiveMs = ms;
                    _mainFmtStr = Fmt(desiredMain); mainTargetFmt = desiredMain;
                    _mainResampling = needRateChange; _mainNoSRC = !needRateChange;
                }
            }

            if (_mainOut == null)
            {
                int ms = BufAligned(_cfg.MainBufMs, false, _defMainMs, 0, _cfg.MainBufMode);
                // Shared: optionally insert internal resampler if enabled
                WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }
                IWaveProvider mainSource = _bufMain;
                if (_cfg.MainForceInternalResamplerInShared && mix != null)
                {
                    var inF = _bufMain.WaveFormat;
                    bool needSRC = inF.SampleRate != mix.SampleRate || inF.Channels != mix.Channels || inF.BitsPerSample != mix.BitsPerSample;
                    if (needSRC)
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, mix);
                        _resMain.ResamplerQuality = _cfg.MainResamplerQuality;
                        mainSource = _resMain;
                    }
                }
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, ms, mainSource, out _mainEvent);
                if (_mainOut == null) { MessageBox.Show("主通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); Cleanup(); return; }
                _mainBufEffectiveMs = ms; try { mainTargetFmt = _outMain.AudioClient.MixFormat; _mainFmtStr = Fmt(mainTargetFmt); } catch { _mainFmtStr = "系统混音"; }
                _mainResampling = (inFmt.SampleRate != (mainTargetFmt != null ? mainTargetFmt.SampleRate : inFmt.SampleRate) || inFmt.Channels != (mainTargetFmt != null ? mainTargetFmt.Channels : inFmt.Channels));
                _mainNoSRC = !_mainResampling;
            }

            // ----- Aux Output -----
            _srcAux = _bufAux; _resAux = null; _auxExclusive = false; _auxEvent = false; _auxBufEffectiveMs = _cfg.AuxBufMs; _auxFmtStr = "-";
            _auxNoSRC = false; _auxResampling = false;
            var desiredAux = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);
            bool isLoopAux = (_inDev.DataFlow == DataFlow.Render) && _inDev.ID == _outAux.ID;
            bool wantExAux = (_cfg.AuxShare == ShareModeOption.Exclusive || _cfg.AuxShare == ShareModeOption.Auto) && !isLoopAux;

            WaveFormat auxTargetFmt = null;

            if (wantExAux && SupportsExclusive(_outAux, desiredAux))
            {
                bool needRateChange = (inFmt.SampleRate != desiredAux.SampleRate) || (inFmt.Channels != desiredAux.Channels);
                if (needRateChange) _srcAux = _resAux = new MediaFoundationResampler(_bufAux, desiredAux) { ResamplerQuality = _cfg.AuxResamplerQuality };
                int ms = BufAligned(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs, _cfg.AuxBufMode);
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, ms, _srcAux, out _auxEvent);
                if (_auxOut != null)
                {
                    _auxExclusive = true;
                    _auxBufEffectiveMs = ms;
                    _auxFmtStr = Fmt(desiredAux); auxTargetFmt = desiredAux;
                    _auxResampling = needRateChange; _auxNoSRC = !needRateChange;
                }
            }

            if (_auxOut == null)
            {
                int ms = BufAligned(_cfg.AuxBufMs, false, _defAuxMs, 0, _cfg.AuxBufMode);
                WaveFormat mix = null; try { mix = _outAux.AudioClient.MixFormat; } catch { }
                IWaveProvider auxSource = _bufAux;
                if (_cfg.AuxForceInternalResamplerInShared && mix != null)
                {
                    var inF = _bufAux.WaveFormat;
                    bool needSRC = inF.SampleRate != mix.SampleRate || inF.Channels != mix.Channels || inF.BitsPerSample != mix.BitsPerSample;
                    if (needSRC)
                    {
                        _resAux = new MediaFoundationResampler(_bufAux, mix);
                        _resAux.ResamplerQuality = _cfg.AuxResamplerQuality;
                        auxSource = _resAux;
                    }
                }
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, ms, auxSource, out _auxEvent);
                if (_auxOut == null) { MessageBox.Show("副通道初始化失败。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error); Cleanup(); return; }
                _auxBufEffectiveMs = ms; try { auxTargetFmt = _outAux.AudioClient.MixFormat; _auxFmtStr = Fmt(auxTargetFmt); } catch { _auxFmtStr = "系统混音"; }
                _auxResampling = (inFmt.SampleRate != (auxTargetFmt != null ? auxTargetFmt.SampleRate : inFmt.SampleRate) || inFmt.Channels != (auxTargetFmt != null ? auxTargetFmt.Channels : inFmt.Channels));
                _auxNoSRC = !_auxResampling;
            }

            _capture.DataAvailable += OnIn;
            _capture.RecordingStopped += OnStopRec;
            try
            {
                _capture.StartRecording();
                _mainOut.Play();
                _auxOut.Play();
                _running = true;
            }
            catch (Exception ex)
            {
                Logger.Crash("Start", ex);
                MessageBox.Show("启动失败：" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void OnIn(object s, WaveInEventArgs e) { if (_bufMain != null) _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded); if (_bufAux != null) _bufAux.AddSamples(e.Buffer, 0, e.BytesRecorded); }
        void OnStopRec(object s, StoppedEventArgs e) { try { if (_bufMain != null) _bufMain.ClearBuffer(); if (_bufAux != null) _bufAux.ClearBuffer(); } catch { } }

        public void Stop()
        {
            try { _capture?.StopRecording(); } catch { }
            try { _mainOut?.Stop(); } catch { }
            try { _auxOut?.Stop(); } catch { }
            Thread.Sleep(20);
            Cleanup();
            _running = false;
        }

        void Cleanup()
        {
            try { if (_capture != null) { _capture.DataAvailable -= OnIn; _capture.RecordingStopped -= OnStopRec; _capture.Dispose(); } } catch { } _capture = null;
            try { _mainOut?.Dispose(); } catch { } _mainOut = null;
            try { _auxOut?.Dispose(); } catch { } _auxOut = null;
            try { _resMain?.Dispose(); } catch { } _resMain = null;
            try { _resAux?.Dispose(); } catch { } _resAux = null;
            _bufMain = null; _bufAux = null;
        }

        // ----- Status -----
        public StatusSnapshot GetStatusSnapshot()
        {
            // Aligned multiples
            double mainMul = 0, auxMul = 0;
            try
            {
                if (_mainBufEffectiveMs > 0) { double baseMs = (_mainExclusive && _minMainMs > 0) ? _minMainMs : _defMainMs; if (baseMs > 0) mainMul = Math.Round(_mainBufEffectiveMs / baseMs, 2); }
                if (_auxBufEffectiveMs > 0) { double baseMs = (_auxExclusive && _minAuxMs > 0) ? _minAuxMs : _defAuxMs; if (baseMs > 0) auxMul = Math.Round(_auxBufEffectiveMs / baseMs, 2); }
            }
            catch { }

            bool mainInternal = _resMain != null;
            bool auxInternal = _resAux != null;
            bool mainMulti = false, auxMulti = false;
            try
            {
                if (!_mainExclusive && mainInternal)
                {
                    WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }
                    if (mix != null && _resMain != null)
                    {
                        var f = _resMain.WaveFormat;
                        if (f.SampleRate != mix.SampleRate || f.Channels != mix.Channels || f.BitsPerSample != mix.BitsPerSample) mainMulti = true;
                    }
                }
                if (!_auxExclusive && auxInternal)
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
                MainDevice = _outMain != null ? _outMain.FriendlyName : "-", AuxDevice = _outAux != null ? _outAux.FriendlyName : "-",
                MainMode = _mainOut != null ? (_mainExclusive ? "独占" : "共享") : "-", AuxMode = _auxOut != null ? (_auxExclusive ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEvent ? "事件" : "轮询") : "-", AuxSync = _auxOut != null ? (_auxEvent ? "事件" : "轮询") : "-",
                MainFormat = _mainOut != null ? _mainFmtStr : "-", AuxFormat = _auxOut != null ? _auxFmtStr : "-",
                MainBufferRequestedMs = _cfg.MainBufMs, AuxBufferRequestedMs = _cfg.AuxBufMs,
                MainBufferMs = _mainOut != null ? _mainBufEffectiveMs : 0, AuxBufferMs = _auxOut != null ? _auxBufEffectiveMs : 0,
                MainDefaultPeriodMs = _defMainMs, MainMinimumPeriodMs = _minMainMs, AuxDefaultPeriodMs = _defAuxMs, AuxMinimumPeriodMs = _minAuxMs,
                MainAlignedMultiple = mainMul, AuxAlignedMultiple = auxMul,
                MainNoSRC = _mainNoSRC, AuxNoSRC = _auxNoSRC, MainResampling = _mainResampling, AuxResampling = _auxResampling,
                MainBufferMultiple = (_mainBufEffectiveMs > 0 && _minMainMs > 0) ? _mainBufEffectiveMs / _minMainMs : 0,
                AuxBufferMultiple = (_auxBufEffectiveMs > 0 && _minAuxMs > 0) ? _auxBufEffectiveMs / _minAuxMs : 0,
                MainInternalResampler = mainInternal, AuxInternalResampler = auxInternal,
                MainMultiSRC = mainMulti, AuxMultiSRC = auxMulti,
                MainInternalResamplerQuality = mainInternal ? _cfg.MainResamplerQuality : -1,
                AuxInternalResamplerQuality = auxInternal ? _cfg.AuxResamplerQuality : -1
            };
        }

        // ----- Helpers -----
        static string Fmt(WaveFormat wf) { return wf == null ? "-" : (wf.SampleRate + "Hz/" + wf.BitsPerSample + "bit/" + wf.Channels + "ch"); }
        MMDevice FindById(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            try { foreach (var d in _mm.EnumerateAudioEndPoints(flow, DeviceState.Active)) if (d.ID == id) return d; } catch { }
            return null;
        }
        void GetPeriods(MMDevice dev, out double defMs, out double minMs) { defMs = 10; minMs = 2; try { var ac = dev.AudioClient; defMs = 10; minMs = 2; } catch { } }
        static bool SupportsExclusive(MMDevice d, WaveFormat f) { try { return d.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, f); } catch { return false; } }
        static int BufAligned(int wantMs, bool exclusive, double defMs, double minMs, BufferAlignMode mode)
        {
            if (exclusive)
            {
                double step = (mode == BufferAlignMode.MinAlign ? (minMs > 0 ? minMs : defMs) : (defMs > 0 ? defMs : minMs));
                double floor = (mode == BufferAlignMode.MinAlign ? step * 3.0 : step * 3.0);
                int ms = (int)Math.Ceiling(Math.Ceiling(wantMs / step) * step);
                if (ms < floor) ms = (int)Math.Ceiling(Math.Ceiling(floor / step) * step);
                return ms;
            }
            else
            {
                double step = (mode == BufferAlignMode.MinAlign ? (minMs > 0 ? minMs : defMs) : (defMs > 0 ? defMs : minMs));
                double floor = (defMs > 0 ? defMs : step) * 2.0;
                int ms = (int)Math.Ceiling(Math.Ceiling(wantMs / step) * step);
                if (ms < floor) ms = (int)Math.Ceiling(Math.Ceiling(floor / step) * step);
                return ms;
            }
        }
        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption pref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false;
            if (pref == SyncModeOption.Polling) return TryOut(dev, mode, false, bufMs, src);
            if (pref == SyncModeOption.Event)
            {
                var w = TryOut(dev, mode, true, bufMs, src);
                if (w != null) { eventUsed = true; return w; }
                return TryOut(dev, mode, false, bufMs, src);
            }
            var w2 = TryOut(dev, mode, true, bufMs, src);
            if (w2 != null) { eventUsed = true; return w2; }
            return TryOut(dev, mode, false, bufMs, src);
        }
        WasapiOut TryOut(MMDevice dev, AudioClientShareMode mode, bool ev, int ms, IWaveProvider src)
        {
            try { var w = new WasapiOut(dev, mode, ev, ms); w.Init(src); return w; } catch { return null; }
        }
        public void Dispose() { Stop(); try { _tray.Visible = false; _tray.Dispose(); _menu.Dispose(); } catch { } }
    }

    // ---------- Input format helper ----------
    public static class InputFormatHelper
    {
        public static WaveFormat BuildWaveFormat(InputFormatStrategy strategy, int customRate, int customBits, int channels)
        {
            switch (strategy)
            {
                case InputFormatStrategy.SystemMix: return null;
                case InputFormatStrategy.Specify24_48000: return CreatePcm24(48000, channels);
                case InputFormatStrategy.Specify24_96000: return CreatePcm24(96000, channels);
                case InputFormatStrategy.Specify24_192000: return CreatePcm24(192000, channels);
                case InputFormatStrategy.Specify32f_48000: return WaveFormat.CreateIeeeFloatWaveFormat(48000, channels);
                case InputFormatStrategy.Specify32f_96000: return WaveFormat.CreateIeeeFloatWaveFormat(96000, channels);
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
    }
}
