using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;

namespace MirrorAudio
{
    enum ShareModeOption { Auto, Exclusive, Shared }
    enum SyncModeOption  { Auto, Event, Polling }

    sealed class AppSettings
    {
        public string InputDeviceId, MainDeviceId, AuxDeviceId;
        public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Auto;
        public SyncModeOption  MainSync  = SyncModeOption.Auto,  AuxSync  = SyncModeOption.Auto;
        public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        public int AuxRate  = 48000,  AuxBits  = 16, AuxBufMs  = 120;
        public bool AutoStart = false, EnableLogging = false;
    }

    sealed class StatusSnapshot
    {
        public bool Running;
        public string InputDevice, InputRole, InputFormat;
        public string MainDevice, MainMode, MainSync, MainFormat;
        public string AuxDevice,  AuxMode,  AuxSync,  AuxFormat;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs;
        public double AuxDefaultPeriodMs,  AuxMinimumPeriodMs;
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
            } catch { }
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

    // 音频图：Capture -> (Resample if needed) -> Main/Aux 两路输出
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
            if (bits == 16) return WaveFormat.CreatePcmWaveFormat(rate, channels);
            // 24/32 位用 IEEE float 以保持兼容性（多数驱动对 float 直通/混音路径最佳）
            return WaveFormat.CreateIeeeFloatWaveFormat(rate, channels);
        }

        static IWaveProvider MaybeResample(IWaveProvider src, WaveFormat dst, int quality)
        {
            if (src.WaveFormat.SampleRate == dst.SampleRate &&
                src.WaveFormat.Channels == dst.Channels &&
                (src.WaveFormat.BitsPerSample == dst.BitsPerSample || dst.Encoding == WaveFormatEncoding.IeeeFloat))
            {
                return src; // 直通
            }
            // 使用 MediaFoundationResampler（质量 1..60）
            var res = new MediaFoundationResampler(src, dst);
            res.ResamplerQuality = quality;
            return res;
        }

        WasapiCapture CreateCapture(MMDevice dev)
        {
            if (dev.DataFlow == DataFlow.Render)
            {
                var loop = new WasapiLoopbackCapture(dev);
                loop.ShareMode = AudioClientShareMode.Shared; // 环回只能共享
                return loop;
            }
            else
            {
                var mic = new WasapiCapture(dev, false, 20); // 20ms 内部缓冲，低开销
                mic.ShareMode = AudioClientShareMode.Shared; // 输入默认共享（更兼容）
                return mic;
            }
        }

        static AudioClientShareMode ResolveShare(ShareModeOption opt, MMDevice dev, bool preferExclusive)
        {
            if (opt == ShareModeOption.Exclusive) return AudioClientShareMode.Exclusive;
            if (opt == ShareModeOption.Shared)   return AudioClientShareMode.Shared;
            return preferExclusive ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared;
        }

        static bool ResolveEvent(SyncModeOption opt, bool preferEvent)
        {
            if (opt == SyncModeOption.Event)   return true;
            if (opt == SyncModeOption.Polling) return false;
            return preferEvent;
        }

        public void Start(out string inFmt, out string mainFmt, out string auxFmt,
                          out bool mainExclusive, out bool auxExclusive,
                          out bool mainEvent, out bool auxEvent,
                          out int effMainBufMs, out int effAuxBufMs)
        {
            // Capture
            _cap = CreateCapture(_inDev);
            _cap.DataAvailable += OnData;
            _cap.RecordingStopped += (s, e) => Logger.Info("Capture stopped: " + (e?.Exception?.Message ?? "OK"));
            _cap.StartRecording();

            _buf = new BufferedWaveProvider(_cap.WaveFormat);
            _buf.DiscardOnBufferOverflow = true;

            inFmt = DescribeFmt(_cap.WaveFormat);

            // Main
            mainExclusive = _exclusiveMain =
                ResolveShare(_cfg.MainShare, _mainDev, true) == AudioClientShareMode.Exclusive;
            mainEvent = _eventMain = ResolveEvent(_cfg.MainSync, true);
            _mainBufMs = _cfg.MainBufMs;

            // Aux
            auxExclusive = _exclusiveAux =
                ResolveShare(_cfg.AuxShare, _auxDev, false) == AudioClientShareMode.Exclusive;
            auxEvent = _eventAux = ResolveEvent(_cfg.AuxSync, true);
            _auxBufMs = _cfg.AuxBufMs;

            // Target formats
            var ch = Math.Max(1, Math.Min( _cap.WaveFormat.Channels, 2)); // 限 1/2 声道
            _targetMain = _exclusiveMain ? MakeTarget(_cfg.MainRate, _cfg.MainBits, ch)
                                         : WaveFormat.CreateIeeeFloatWaveFormat(_cap.WaveFormat.SampleRate, ch);
            _targetAux  = _exclusiveAux  ? MakeTarget(_cfg.AuxRate,  _cfg.AuxBits,  ch)
                                         : WaveFormat.CreateIeeeFloatWaveFormat(_cap.WaveFormat.SampleRate, ch);

            // Resampler：主=50，副=40
            var src = (IWaveProvider)_buf;
            _toMain = MaybeResample(src, _targetMain, 50);
            _toAux  = MaybeResample(src, _targetAux,  40);

            // Outputs
            _outMain = new WasapiOut(_mainDev, _exclusiveMain ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared,
                                     _eventMain, _mainBufMs);
            _outMain.Init(_toMain);
            _outMain.Play();

            _outAux = new WasapiOut(_auxDev, _exclusiveAux ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared,
                                    _eventAux, _auxBufMs);
            _outAux.Init(_toAux);
            _outAux.Play();

            mainFmt = DescribeFmt(_targetMain);
            auxFmt  = DescribeFmt(_targetAux);
            effMainBufMs = _mainBufMs;
            effAuxBufMs  = _auxBufMs;
        }

        void OnData(object s, WaveInEventArgs e)
        {
            if (e.BytesRecorded > 0) _buf.AddSamples(e.Buffer, 0, e.BytesRecorded);
        }

        public void StopSafe()
        {
            try { if (_outMain != null) _outMain.Stop(); } catch { }
            try { if (_outAux  != null) _outAux.Stop();  } catch { }
            try { if (_cap     != null) _cap.StopRecording(); } catch { }

            Dispose();
        }

        public void Dispose()
        {
            try { if (_outMain != null) _outMain.Dispose(); } catch { }
            try { if (_outAux  != null) _outAux.Dispose();  } catch { }
            try { if (_cap     != null) _cap.Dispose();     } catch { }
            _outMain = null; _outAux = null; _cap = null; _buf = null;
        }

        static string DescribeFmt(WaveFormat f)
        {
            if (f == null) return "-";
            var bits = f.BitsPerSample;
            var enc = f.Encoding == WaveFormatEncoding.IeeeFloat ? "float" : (bits + "bit");
            return f.SampleRate + " Hz, " + enc + ", " + f.Channels + "ch";
        }
    }

    // 托盘与设置
    sealed class TrayApp : Form, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();
        readonly System.Windows.Forms.Timer _debounce = new System.Windows.Forms.Timer() { Interval = 300 };
        AppSettings _cfg = new AppSettings();

        // 运行态
        AudioGraph _graph;
        bool _running, _mainIsExclusive, _mainEventSyncUsed, _auxIsExclusive, _auxEventSyncUsed;
        int _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-", _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(4);

        public TrayApp()
        {
            Logger.Enabled = _cfg.EnableLogging;
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            // 托盘图标：优先用 Assets\MirrorAudio.ico
            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                if (File.Exists(icoPath)) _tray.Icon = new Icon(icoPath);
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop  = new ToolStripMenuItem("停止(&T)",     null, (s, e) => Stop());
            var miSet   = new ToolStripMenuItem("设置(&G)...",  null, (s, e) => OnSettings());
            var miLog   = new ToolStripMenuItem("打开日志目录",   null, (s, e) => Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)",     null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s, e) => { _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        void EnsureAutoStart(bool on)
        {
            try
            {
                var rk = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                if (on) rk.SetValue("MirrorAudio", Application.ExecutablePath);
                else rk.DeleteValue("MirrorAudio", false);
            } catch { }
        }

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    _cfg = f.Result;
                    Logger.Enabled = _cfg.EnableLogging;
                    EnsureAutoStart(_cfg.AutoStart);
                    StartOrRestart();
                }
            }
        }

        void StartOrRestart()
        {
            Stop();

            try
            {
                var inDev   = !string.IsNullOrEmpty(_cfg.InputDeviceId) ? _mm.GetDevice(_cfg.InputDeviceId) : GetDefaultInput();
                var mainDev = _mm.GetDevice(_cfg.MainDeviceId ?? _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).ID);
                var auxDev  = _mm.GetDevice(_cfg.AuxDeviceId  ?? _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console).ID);

                _inDevName = inDev.FriendlyName;
                _inRoleStr = inDev.DataFlow == DataFlow.Render ? "环回" : "录音";

                _graph = new AudioGraph(_cfg, inDev, mainDev, auxDev);
                string inFmt, mainFmt, auxFmt;
                _graph.Start(out inFmt, out mainFmt, out auxFmt,
                             out _mainIsExclusive, out _auxIsExclusive,
                             out _mainEventSyncUsed, out _auxEventSyncUsed,
                             out _mainBufEffectiveMs, out _auxBufEffectiveMs);
                _inFmtStr = inFmt; _mainFmtStr = mainFmt; _auxFmtStr = auxFmt;

                var pm = GetPeriods(mainDev.ID);
                _defMainMs = pm.Item1; _minMainMs = pm.Item2;
                pm = GetPeriods(auxDev.ID);
                _defAuxMs = pm.Item1; _minAuxMs = pm.Item2;

                _running = true;
                Logger.Info("Started.");
            }
            catch (Exception ex)
            {
                _running = false;
                Logger.Info("Start failed: " + ex.Message);
                MessageBox.Show("启动失败：\n" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void Stop()
        {
            _running = false;
            if (_graph != null) { _graph.StopSafe(); _graph = null; }
        }

        StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot();
            s.Running = _running;
            s.InputDevice = _inDevName;
            s.InputRole   = _inRoleStr;
            s.InputFormat = _inFmtStr;

            s.MainDevice = DescribeDevice(_cfg.MainDeviceId);
            s.MainMode   = _mainIsExclusive ? "独占" : "共享";
            s.MainSync   = _mainEventSyncUsed ? "事件" : "轮询";
            s.MainFormat = _mainFmtStr;
            s.MainBufferMs = _mainBufEffectiveMs;
            s.MainDefaultPeriodMs = _defMainMs;
            s.MainMinimumPeriodMs = _minMainMs;

            s.AuxDevice = DescribeDevice(_cfg.AuxDeviceId);
            s.AuxMode   = _auxIsExclusive ? "独占" : "共享";
            s.AuxSync   = _auxEventSyncUsed ? "事件" : "轮询";
            s.AuxFormat = _auxFmtStr;
            s.AuxBufferMs = _auxBufEffectiveMs;
            s.AuxDefaultPeriodMs = _defAuxMs;
            s.AuxMinimumPeriodMs = _minAuxMs;
            return s;
        }

        string DescribeDevice(string id)
        {
            try { if (!string.IsNullOrEmpty(id)) return _mm.GetDevice(id).FriendlyName; } catch { }
            return "-";
        }

        MMDevice GetDefaultInput()
        {
            try { return _mm.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications); } catch { }
            try { return _mm.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia); } catch { }
            // 兜底：随便取一个 Render 做环回
            return _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        }

        // —— 设备周期：反射尝试一次，失败退 10/2 ms —— //
        Tuple<double, double> GetPeriods(string devId)
        {
            Tuple<double, double> t;
            if (_periodCache.TryGetValue(devId ?? "", out t)) return t;
            double defMs = 10, minMs = 2;
            try
            {
                var dev = _mm.GetDevice(devId);
                var ac = dev.AudioClient;
                // 某些 NAudio 版本上公开 GetDevicePeriod；若不可用则走异常
                long def, min;
                try
                {
                    ac.GetDevicePeriod(out def, out min);
                    defMs = def / 10000.0; // 100ns -> ms
                    minMs = min / 10000.0;
                }
                catch
                {
                    // 反射兜底（若无则维持 10/2）
                    var m = ac.GetType().GetMethod("GetDevicePeriod");
                    if (m != null)
                    {
                        object[] args = new object[] { 0L, 0L };
                        m.Invoke(ac, args);
                        defMs = (long)args[0] / 10000.0;
                        minMs = (long)args[1] / 10000.0;
                    }
                }
            } catch { }
            t = Tuple.Create(defMs, minMs);
            _periodCache[devId ?? ""] = t;
            return t;
        }

        // —— IMMNotificationClient：设备热插拔事件驱动 + 去抖 —— //
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
