using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;

namespace MirrorAudio
{
    enum ShareModeOption { Auto, Exclusive, Shared }
    enum SyncModeOption { Auto, Event, Polling }

    sealed class AppSettings
    {
        public string InputDeviceId, MainDeviceId, AuxDeviceId;
        public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Auto;
        public SyncModeOption MainSync = SyncModeOption.Auto, AuxSync = SyncModeOption.Auto;
        public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        public int AuxRate = 48000, AuxBits = 16, AuxBufMs = 120;
        public bool AutoStart = false, EnableLogging = false;
    }

    sealed class StatusSnapshot
    {
        public bool Running;
        public string InputDevice, InputRole, InputFormat;
        public string MainDevice, MainMode, MainSync, MainFormat;
        public string AuxDevice, AuxMode, AuxSync, AuxFormat;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs;
        public double AuxDefaultPeriodMs, AuxMinimumPeriodMs;
    }

    static class Logger
    {
        static bool _enabled;
        public static bool Enabled { get { return _enabled; } set { _enabled = value; } }
        public static void Info(string s)
        {
            if (!_enabled) return;
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "MirrorAudio.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff ") + s + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApp());
        }
    }

    // —— 音频图 —— //
    sealed class AudioGraph : IDisposable
    {
        readonly AppSettings _cfg;
        readonly MMDevice _inDev, _mainDev, _auxDev;
        WasapiCapture _cap;
        IWavePlayer _outMain, _outAux;
        BufferedWaveProvider _buf;
        IWaveProvider _toMain, _toAux;
        WaveFormat _targetMain, _targetAux;
        int _mainBufMs, _auxBufMs;
        bool _eventMain, _eventAux;
        bool _exclusiveMain, _exclusiveAux;

        public AudioGraph(AppSettings cfg, MMDevice inDev, MMDevice mainDev, MMDevice auxDev)
        {
            _cfg = cfg; _inDev = inDev; _mainDev = mainDev; _auxDev = auxDev;
        }

        static WaveFormat MakeTarget(int rate, int bits, int channels)
        {
            if (bits == 16) return new WaveFormat(rate, bits, channels);
            return WaveFormat.CreateIeeeFloatWaveFormat(rate, channels);
        }

        static IWaveProvider MaybeResample(IWaveProvider src, WaveFormat dst, int quality)
        {
            if (src.WaveFormat.SampleRate == dst.SampleRate &&
                src.WaveFormat.Channels == dst.Channels &&
                (src.WaveFormat.BitsPerSample == dst.BitsPerSample || dst.Encoding == WaveFormatEncoding.IeeeFloat))
            {
                return src;
            }
            var res = new MediaFoundationResampler(src, dst);
            res.ResamplerQuality = quality;
            return res;
        }

        WasapiCapture CreateCapture(MMDevice dev)
        {
            if (dev.DataFlow == DataFlow.Render)
            {
                var loop = new WasapiLoopbackCapture(dev);
                loop.ShareMode = AudioClientShareMode.Shared;
                return loop;
            }
            else
            {
                var mic = new WasapiCapture(dev, false, 20);
                mic.ShareMode = AudioClientShareMode.Shared;
                return mic;
            }
        }

        static AudioClientShareMode ResolveShare(ShareModeOption opt, bool preferExclusive)
        {
            if (opt == ShareModeOption.Exclusive) return AudioClientShareMode.Exclusive;
            if (opt == ShareModeOption.Shared) return AudioClientShareMode.Shared;
            return preferExclusive ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared;
        }

        static bool ResolveEvent(SyncModeOption opt, bool preferEvent)
        {
            if (opt == SyncModeOption.Event) return true;
            if (opt == SyncModeOption.Polling) return false;
            return preferEvent;
        }

        public void Start(out string inFmt, out string mainFmt, out string auxFmt,
                          out bool mainExclusive, out bool auxExclusive,
                          out bool mainEvent, out bool auxEvent,
                          out int effMainBufMs, out int effAuxBufMs)
        {
            _cap = CreateCapture(_inDev);
            _cap.DataAvailable += OnData;
            _cap.StartRecording();

            _buf = new BufferedWaveProvider(_cap.WaveFormat) { DiscardOnBufferOverflow = true };
            inFmt = DescribeFmt(_cap.WaveFormat);

            mainExclusive = _exclusiveMain = ResolveShare(_cfg.MainShare, true) == AudioClientShareMode.Exclusive;
            mainEvent = _eventMain = ResolveEvent(_cfg.MainSync, true);
            _mainBufMs = _cfg.MainBufMs;

            auxExclusive = _exclusiveAux = ResolveShare(_cfg.AuxShare, false) == AudioClientShareMode.Exclusive;
            auxEvent = _eventAux = ResolveEvent(_cfg.AuxSync, true);
            _auxBufMs = _cfg.AuxBufMs;

            var ch = Math.Max(1, Math.Min(_cap.WaveFormat.Channels, 2));
            _targetMain = _exclusiveMain ? MakeTarget(_cfg.MainRate, _cfg.MainBits, ch)
                                         : WaveFormat.CreateIeeeFloatWaveFormat(_cap.WaveFormat.SampleRate, ch);
            _targetAux = _exclusiveAux ? MakeTarget(_cfg.AuxRate, _cfg.AuxBits, ch)
                                       : WaveFormat.CreateIeeeFloatWaveFormat(_cap.WaveFormat.SampleRate, ch);

            _toMain = MaybeResample(_buf, _targetMain, 50);
            _toAux = MaybeResample(_buf, _targetAux, 40);

            _outMain = new WasapiOut(_mainDev, _exclusiveMain ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared,
                                     _eventMain, _mainBufMs);
            _outMain.Init(_toMain); _outMain.Play();

            _outAux = new WasapiOut(_auxDev, _exclusiveAux ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared,
                                    _eventAux, _auxBufMs);
            _outAux.Init(_toAux); _outAux.Play();

            mainFmt = DescribeFmt(_targetMain);
            auxFmt = DescribeFmt(_targetAux);
            effMainBufMs = _mainBufMs;
            effAuxBufMs = _auxBufMs;
        }

        void OnData(object s, WaveInEventArgs e) => _buf.AddSamples(e.Buffer, 0, e.BytesRecorded);

        public void StopSafe()
        {
            try { _outMain?.Stop(); } catch { }
            try { _outAux?.Stop(); } catch { }
            try { _cap?.StopRecording(); } catch { }
            Dispose();
        }

        public void Dispose()
        {
            try { _outMain?.Dispose(); } catch { }
            try { _outAux?.Dispose(); } catch { }
            try { _cap?.Dispose(); } catch { }
            _outMain = null; _outAux = null; _cap = null; _buf = null;
        }

        static string DescribeFmt(WaveFormat f)
        {
            if (f == null) return "-";
            var enc = f.Encoding == WaveFormatEncoding.IeeeFloat ? "float" : (f.BitsPerSample + "bit");
            return $"{f.SampleRate} Hz, {enc}, {f.Channels}ch";
        }
    }

    sealed class TrayApp : Form, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();
        readonly System.Windows.Forms.Timer _debounce = new System.Windows.Forms.Timer() { Interval = 300 };
        AppSettings _cfg = new AppSettings();
        AudioGraph _graph;
        bool _running, _mainIsExclusive, _mainEventSyncUsed, _auxIsExclusive, _auxEventSyncUsed;
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;
        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(4);

        public TrayApp()
        {
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }
            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                _tray.Icon = File.Exists(icoPath) ? new Icon(icoPath) : Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启", null, (s, e) => StartOrRestart());
            var miStop = new ToolStripMenuItem("停止", null, (s, e) => Stop());
            var miSet = new ToolStripMenuItem("设置...", null, (s, e) => OnSettings());
            var miExit = new ToolStripMenuItem("退出", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s, e) => { _debounce.Stop(); StartOrRestart(); };
            StartOrRestart();
        }

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    _cfg = f.Result;
                    StartOrRestart();
                }
            }
        }

        void StartOrRestart()
        {
            Stop();
            try
            {
                var inDev = !string.IsNullOrEmpty(_cfg.InputDeviceId) ? _mm.GetDevice(_cfg.InputDeviceId) : _mm.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
                var mainDev = _mm.GetDevice(_cfg.MainDeviceId ?? _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).ID);
                var auxDev = _mm.GetDevice(_cfg.AuxDeviceId ?? _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console).ID);

                _inDevName = inDev.FriendlyName;
                _inRoleStr = inDev.DataFlow == DataFlow.Render ? "环回" : "录音";

                _graph = new AudioGraph(_cfg, inDev, mainDev, auxDev);
                _graph.Start(out _inFmtStr, out _mainFmtStr, out _auxFmtStr,
                             out _mainIsExclusive, out _auxIsExclusive,
                             out _mainEventSyncUsed, out _auxEventSyncUsed,
                             out _mainBufEffectiveMs, out _auxBufEffectiveMs);

                var pm = GetPeriods(mainDev.ID);
                _defMainMs = pm.Item1; _minMainMs = pm.Item2;
                pm = GetPeriods(auxDev.ID);
                _defAuxMs = pm.Item1; _minAuxMs = pm.Item2;
                _running = true;
            }
            catch { _running = false; Stop(); }
        }

        void Stop()
        {
            _running = false;
            _graph?.StopSafe();
            _graph = null;
        }

        StatusSnapshot GetStatusSnapshot() => new StatusSnapshot
        {
            Running = _running,
            InputDevice = _inDevName,
            InputRole = _inRoleStr,
            InputFormat = _inFmtStr,
            MainDevice = SafeName(_cfg.MainDeviceId),
            MainMode = _mainIsExclusive ? "独占" : "共享",
            MainSync = _mainEventSyncUsed ? "事件" : "轮询",
            MainFormat = _mainFmtStr,
            MainBufferMs = _mainBufEffectiveMs,
            MainDefaultPeriodMs = _defMainMs,
            MainMinimumPeriodMs = _minMainMs,
            AuxDevice = SafeName(_cfg.AuxDeviceId),
            AuxMode = _auxIsExclusive ? "独占" : "共享",
            AuxSync = _auxEventSyncUsed ? "事件" : "轮询",
            AuxFormat = _auxFmtStr,
            AuxBufferMs = _auxBufEffectiveMs,
            AuxDefaultPeriodMs = _defAuxMs,
            AuxMinimumPeriodMs = _minAuxMs
        };

        string SafeName(string id)
        {
            try { return string.IsNullOrEmpty(id) ? "-" : _mm.GetDevice(id).FriendlyName; }
            catch { return "-"; }
        }

        // —— 周期兼容 —— //
        Tuple<double, double> GetPeriods(string devId)
        {
            if (_periodCache.TryGetValue(devId ?? "", out var t)) return t;
            double defMs = 10, minMs = 2;
            try
            {
                var dev = _mm.GetDevice(devId);
                var ac = dev.AudioClient;
                var tp = ac.GetType();
                var defProp = tp.GetProperty("DefaultDevicePeriod");
                var minProp = tp.GetProperty("MinimumDevicePeriod");
                if (defProp != null && minProp != null)
                {
                    defMs = ToMs(defProp.GetValue(ac));
                    minMs = ToMs(minProp.GetValue(ac));
                }
                else
                {
                    var m = tp.GetMethod("GetDevicePeriod");
                    if (m != null)
                    {
                        object[] args = new object[] { 0L, 0L };
                        m.Invoke(ac, args);
                        defMs = ((long)args[0]) / 10000.0;
                        minMs = ((long)args[1]) / 10000.0;
                    }
                }
            }
            catch { }
            t = Tuple.Create(defMs, minMs);
            _periodCache[devId ?? ""] = t;
            return t;
        }

        static double ToMs(object v)
        {
            if (v == null) return 0;
            if (v is TimeSpan ts) return ts.TotalMilliseconds;
            if (v is long l) return l / 10000.0;
            try { return Convert.ToInt64(v) / 10000.0; } catch { return 0; }
        }

        // IMMNotificationClient
        public void OnDeviceStateChanged(string deviceId, DeviceState newState) { _debounce.Stop(); _debounce.Start(); }
        public void OnDeviceAdded(string pwstrDeviceId) { _debounce.Stop(); _debounce.Start(); }
        public void OnDeviceRemoved(string deviceId) { _debounce.Stop(); _debounce.Start(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId) { _debounce.Stop(); _debounce.Start(); }
        public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key) { }
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try { _mm.UnregisterEndpointNotificationCallback(this); } catch { }
            Stop();
            base.OnFormClosed(e);
        }
    }
}
