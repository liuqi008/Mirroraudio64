using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;

namespace MirrorAudio
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try { Logger.Init(); } catch { }
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Log("unhandled: " + e.ExceptionObject);
            Application.ThreadException += (s, e) => Logger.Log("ui-ex: " + e.Exception);

            Application.Run(new TrayApp());
        }
    }

    // ============== 配置 & 状态 ==============

    public enum ShareModeOption { Auto, Exclusive, Shared }
    public enum SyncModeOption { Auto, Event, Polling }
    public enum PathType { None, PassthroughExclusive, PassthroughSharedMix, Resampled }

    public sealed class AppSettings
    {
        public string InputDeviceId, MainDeviceId, AuxDeviceId;
        public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Auto;
        public SyncModeOption MainSync = SyncModeOption.Auto, AuxSync = SyncModeOption.Auto;
        public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        public int AuxRate = 48000, AuxBits = 16, AuxBufMs = 160;
        public bool AutoStart = false, EnableLogging = false;
        public int AuxResamplerQuality = 40; // 30/40/50
    }

    public sealed class StatusSnapshot
    {
        public bool Running;
        public string InputDevice, InputRole, InputFormat;
        public string MainDevice, MainMode, MainSync, MainFormat;
        public string AuxDevice, AuxMode, AuxSync, AuxFormat;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs, AuxDefaultPeriodMs, AuxMinimumPeriodMs;

        public bool MainPassthrough, AuxPassthrough;
        public string MainPassDesc, AuxPassDesc; // “独占直通/共享混音直通/重采样”
        public int AuxQuality;                   // 30/40/50
    }

    // ============== 托盘主类 ==============

    public sealed class TrayApp : Form, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly System.Windows.Forms.Timer _debounce = new System.Windows.Forms.Timer() { Interval = 600 };

        readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();
        AppSettings _cfg = ConfigStore.Load();
        // 设备周期缓存（设备ID->(default,min)）
        readonly Dictionary<string, Tuple<double, double>> _periodCache = new Dictionary<string, Tuple<double, double>>(8);

        // 输入链
        WasapiCapture _cap;
        BufferedWaveProvider _bufMain, _bufAux;

        // 输出链 - 主
        MMDevice _outMain;
        WasapiOut _mainOut;
        MediaFoundationResampler _resMain;
        bool _mainIsExclusive, _mainEventSyncUsed;
        int _mainBufEffectiveMs;
        PathType _mainPath = PathType.None;

        // 输出链 - 副
        MMDevice _outAux;
        WasapiOut _auxOut;
        MediaFoundationResampler _resAux;
        bool _auxIsExclusive, _auxEventSyncUsed;
        int _auxBufEffectiveMs;
        PathType _auxPath = PathType.None;

        // 状态展示字段
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-";
        string _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        public TrayApp()
        {
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            // 托盘图标：优先 Assets\MirrorAudio.ico
            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                if (File.Exists(icoPath)) _tray.Icon = new Icon(icoPath);
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            } catch { _tray.Icon = SystemIcons.Application; }

            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var miStop  = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, (s, e) => OnSettings());
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s, e) => Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[] { miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s, e) => { _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            Logger.Enabled = _cfg.EnableLogging;
            StartOrRestart();
        }

        // ============== 设置窗口 ==============

        void OnSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog(this) == DialogResult.OK && f.Result != null)
                {
                    var old = _cfg; _cfg = f.Result;
                    Logger.Enabled = _cfg.EnableLogging;
                    EnsureAutoStart(_cfg.AutoStart);
                    if (!ConfigStore.Equals(old, _cfg)) ConfigStore.Save(_cfg);
                    // 重启音频
                    _debounce.Stop(); _debounce.Start();
                }
            }
        }

        // ============== 启停链路 ==============

        void StartOrRestart()
        {
            Stop();
            try
            {
                // 1) 输入设备
                var capDev = PickInput(_cfg.InputDeviceId);
                if (capDev == null) { MessageBox.Show("未找到输入设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                _inDevName = capDev.FriendlyName;
                bool isLoopback = capDev.DataFlow == DataFlow.Render;

                _cap = isLoopback ? (WasapiCapture)new WasapiLoopbackCapture(capDev) : new WasapiCapture(capDev);
                _cap.ShareMode = AudioClientShareMode.Shared;
                _cap.DataAvailable += OnData;
                _cap.RecordingStopped += (s, e) => Logger.Log("capture stopped: " + e.Exception);

                var inFmt = _cap.WaveFormat;
                _inRoleStr = isLoopback ? "环回" : "录音";
                _inFmtStr = Fmt(inFmt);

                // 2) 输出设备
                _outMain = PickRender(_cfg.MainDeviceId);
                _outAux  = PickRender(_cfg.AuxDeviceId);
                if (_outMain == null || _outAux == null) { MessageBox.Show("请在设置中选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                // 设备周期（一次并缓存）
                var perM = GetOrCachePeriod(_outMain); _defMainMs = perM.Item1; _minMainMs = perM.Item2;
                var perA = GetOrCachePeriod(_outAux);  _defAuxMs  = perA.Item1; _minAuxMs  = perA.Item2;

                // 3) 环形缓存（更紧凑）
                _bufMain = new BufferedWaveProvider(inFmt)
                {
                    DiscardOnBufferOverflow = true, ReadFully = true,
                    BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 6, 96))
                };
                _bufAux = new BufferedWaveProvider(inFmt)
                {
                    DiscardOnBufferOverflow = true, ReadFully = true,
                    BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs * 4, 120))
                };

                // 4) 主通道：独占优先、能直通就直通；否则共享并尝试混音直通；否则重采样
                InitMainPath(inFmt);

                // 5) 副通道：同理，但重采样质量来自用户（30/40/50）
                InitAuxPath(inFmt);

                // 6) 开始
                _cap.StartRecording();
                if (_mainOut != null) _mainOut.Play();
                if (_auxOut  != null) _auxOut.Play();
            }
            catch (Exception ex)
            {
                Logger.Log("Start failed: " + ex);
                MessageBox.Show("启动失败：" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void Stop()
        {
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut .Stop(); } catch { }
            try { if (_cap != null) _cap.StopRecording(); } catch { }

            DisposeSafe(ref _resMain);
            DisposeSafe(ref _resAux);
            DisposeSafe(ref _mainOut);
            DisposeSafe(ref _auxOut);
            DisposeSafe(ref _cap);
            _mainPath = PathType.None; _auxPath = PathType.None;
        }

        // ============== 主/副链初始化 ==============

        void InitMainPath(WaveFormat inFmt)
        {
            _mainIsExclusive = false; _mainEventSyncUsed = false;
            _mainPath = PathType.None;

            bool wantExclusive = _cfg.MainShare == ShareModeOption.Exclusive || (_cfg.MainShare == ShareModeOption.Auto && PreferExclusive());
            var desiredMain = WaveFormatFromUser(_cfg.MainRate, _cfg.MainBits, inFmt.Channels, preferFloat: false);

            if (wantExclusive && SupportsExclusive(_outMain, inFmt))
            {
                // 独占直通（输入格式）
                int ms = Buf(_cfg.MainBufMs, true, _defMainMs, _minMainMs);
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _bufMain, out _mainEventSyncUsed, inFmt);
                if (_mainOut != null)
                {
                    _mainIsExclusive = true; _mainBufEffectiveMs = ms; _mainFmtStr = Fmt(inFmt); _mainPath = PathType.PassthroughExclusive;
                    return;
                }
            }

            if (wantExclusive && desiredMain != null && SupportsExclusive(_outMain, desiredMain))
            {
                // 独占直通（用户目标）
                int ms = Buf(_cfg.MainBufMs, true, _defMainMs, _minMainMs);
                _mainOut = CreateOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _bufMain, out _mainEventSyncUsed, desiredMain);
                if (_mainOut != null)
                {
                    _mainIsExclusive = true; _mainBufEffectiveMs = ms; _mainFmtStr = Fmt(desiredMain); _mainPath = PathType.PassthroughExclusive;
                    return;
                }
            }

            // 共享：混音直通优先，否则重采样到混音格式
            int msS = Buf(_cfg.MainBufMs, false, _defMainMs);
            WaveFormat mix = null; try { mix = _outMain.AudioClient.MixFormat; } catch { }
            IWaveProvider src = _bufMain;

            if (mix != null && !Eq(mix, _bufMain.WaveFormat))
            {
                _resMain = new MediaFoundationResampler(_bufMain, mix) { ResamplerQuality = 50 };
                src = _resMain; _mainPath = PathType.Resampled;
            }
            else
            {
                _mainPath = PathType.PassthroughSharedMix;
            }

            _mainOut = CreateOut(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, msS, src, out _mainEventSyncUsed);
            if (_mainOut != null) { _mainBufEffectiveMs = msS; _mainFmtStr = Fmt(mix ?? _bufMain.WaveFormat); }
        }

        void InitAuxPath(WaveFormat inFmt)
        {
            _auxIsExclusive = false; _auxEventSyncUsed = false;
            _auxPath = PathType.None;

            bool wantExclusive = _cfg.AuxShare == ShareModeOption.Exclusive || (_cfg.AuxShare == ShareModeOption.Auto && PreferExclusive());
            var desiredAux = WaveFormatFromUser(_cfg.AuxRate, _cfg.AuxBits, inFmt.Channels, preferFloat: false);

            if (wantExclusive && SupportsExclusive(_outAux, inFmt))
            {
                int ms = Buf(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs);
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, ms, _bufAux, out _auxEventSyncUsed, inFmt);
                if (_auxOut != null)
                {
                    _auxIsExclusive = true; _auxBufEffectiveMs = ms; _auxFmtStr = Fmt(inFmt); _auxPath = PathType.PassthroughExclusive;
                    return;
                }
            }

            if (wantExclusive && desiredAux != null && SupportsExclusive(_outAux, desiredAux))
            {
                int ms = Buf(_cfg.AuxBufMs, true, _defAuxMs, _minAuxMs);
                _auxOut = CreateOut(_outAux, AudioClientShareMode.Exclusive, _cfg.AuxSync, ms, _bufAux, out _auxEventSyncUsed, desiredAux);
                if (_auxOut != null)
                {
                    _auxIsExclusive = true; _auxBufEffectiveMs = ms; _auxFmtStr = Fmt(desiredAux); _auxPath = PathType.PassthroughExclusive;
                    return;
                }
            }

            // 共享：混音直通优先，否则重采样到混音格式（质量=手动 30/40/50）
            int msS = Buf(_cfg.AuxBufMs, false, _defAuxMs);
            WaveFormat mix = null; try { mix = _outAux.AudioClient.MixFormat; } catch { }
            IWaveProvider src = _bufAux;

            if (mix != null && !Eq(mix, _bufAux.WwaveFormat))
            {
                int q = Clamp(_cfg.AuxResamplerQuality, 30, 50);
                _resAux = new MediaFoundationResampler(_bufAux, mix) { ResamplerQuality = q };
                src = _resAux; _auxPath = PathType.Resampled;
            }
            else
            {
                _auxPath = PathType.PassthroughSharedMix;
            }

            _auxOut = CreateOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, msS, src, out _auxEventSyncUsed);
            if (_auxOut != null) { _auxBufEffectiveMs = msS; _auxFmtStr = Fmt(mix ?? _bufAux.WaveFormat); }
        }

        // ============== 数据回调 ==============

        void OnData(object s, WaveInEventArgs e)
        {
            try
            {
                if (_bufMain != null) _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded);
                if (_bufAux  != null) _bufAux .AddSamples(e.Buffer, 0, e.BytesRecorded);
            }
            catch (Exception ex) { Logger.Log("OnData: " + ex.Message); }
        }

        // ============== 工具：创建输出、判等、周期 ==============

        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption sync, int ms, IWaveProvider src, out bool eventUsed, WaveFormat forceFormat = null)
        {
            eventUsed = false;
            try
            {
                var useEvent = (sync == SyncModeOption.Event) || (sync == SyncModeOption.Auto && PreferEvent());
                var wo = new WasapiOut(dev, mode, useEvent ? 0 : 0, ms);
                eventUsed = useEvent;

                if (forceFormat != null)
                {
                    // 对于独占强制格式的情况，用一个格式转换包装器（在直通时不需要）
                    if (!Eq(src.WaveFormat, forceFormat))
                        src = new MediaFoundationResampler(src, forceFormat) { ResamplerQuality = 50 };
                }
                wo.Init(src);
                return wo;
            }
            catch (Exception ex) { Logger.Log("CreateOut: " + ex.Message); return null; }
        }

        static bool PreferExclusive() { return true; } // 保持“独占优先”的默认倾向
        static bool PreferEvent() { return true; }     // 事件优先；用户可改为轮询

        static int Buf(int wantMs, bool isExclusive, double defMs, double minMs = 2)
        {
            // 独占：对齐到设备最小周期倍数；共享：直接用 want
            int ms = Math.Max(4, wantMs);
            if (isExclusive)
            {
                // 简化：向上取整到 >= wantMs 的最小周期倍数
                var m = Math.Max(1, (int)Math.Round(minMs));
                int k = (int)Math.Ceiling(ms / (double)m);
                ms = Math.Max((int)(k * m), (int)Math.Ceiling(minMs));
            }
            return ms;
        }

        WaveFormat WaveFormatFromUser(int rate, int bits, int ch, bool preferFloat)
        {
            try
            {
                if (preferFloat && bits >= 32) return WaveFormat.CreateIeeeFloatWaveFormat(rate, ch);
                if (bits == 24)
                    return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, ch, rate * ch * 3, ch * 3, 24);
                if (bits == 32 && preferFloat)
                    return WaveFormat.CreateIeeeFloatWaveFormat(rate, ch);
                if (bits == 32)
                    return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, ch, rate * ch * 4, ch * 4, 32);
                return new WaveFormat(rate, bits, ch);
            } catch { return null; }
        }

        static bool Eq(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels && a.Encoding == b.Encoding;
        }

        Tuple<double, double> GetOrCachePeriod(MMDevice dev)
        {
            if (dev == null) return Tuple.Create(10.0, 2.0);
            Tuple<double, double> t;
            if (_periodCache.TryGetValue(dev.ID, out t)) return t;
            double defMs = 10, minMs = 2;
            try
            {
                var ac = dev.AudioClient;
                long def100ns, min100ns;
                ac.GetDevicePeriod(out def100ns, out min100ns);
                defMs = def100ns / 10000.0; minMs = min100ns / 10000.0;
            }
            catch { }
            t = Tuple.Create(defMs, minMs);
            _periodCache[dev.ID] = t;
            return t;
        }

        // ============== 设备选择/热插拔 ==============

        MMDevice PickInput(string id)
        {
            try
            {
                if (!string.IsNullOrEmpty(id))
                {
                    var d = _mm.GetDevice(id);
                    if (d.State == DeviceState.Active && (d.DataFlow == DataFlow.Capture || d.DataFlow == DataFlow.Render)) return d;
                }
            } catch { }
            // fallback：默认通信设备（优先）
            try { return _mm.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console); } catch { }
            try { return _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console); } catch { }
            return null;
        }

        MMDevice PickRender(string id)
        {
            try
            {
                if (!string.IsNullOrEmpty(id))
                {
                    var d = _mm.GetDevice(id);
                    if (d.State == DeviceState.Active && d.DataFlow == DataFlow.Render) return d;
                }
            } catch { }
            try { return _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia); } catch { }
            return null;
        }

        bool SupportsExclusive(MMDevice dev, WaveFormat fmt)
        {
            try
            {
                var ac = dev.AudioClient;
                var hr = ac.IsFormatSupported(AudioClientShareMode.Exclusive, fmt);
                return hr;
            } catch { return false; }
        }

        static void DisposeSafe<T>(ref T o) where T : class, IDisposable
        {
            try { if (o != null) o.Dispose(); } catch { } finally { o = null; }
        }

        // ============== 状态 & 设置 ==============

        public StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot();
            try
            {
                s.Running = _cap != null && _mainOut != null && _auxOut != null;
                s.InputDevice = _inDevName; s.InputRole = _inRoleStr; s.InputFormat = _inFmtStr;

                s.MainDevice = _outMain != null ? _outMain.FriendlyName : null;
                s.AuxDevice  = _outAux  != null ? _outAux .FriendlyName : null;

                s.MainMode = _mainIsExclusive ? "独占" : "共享";
                s.AuxMode  = _auxIsExclusive  ? "独占" : "共享";

                s.MainSync = _mainEventSyncUsed ? "事件" : "轮询";
                s.AuxSync  = _auxEventSyncUsed  ? "事件" : "轮询";

                s.MainFormat = _mainFmtStr; s.AuxFormat = _auxFmtStr;
                s.MainBufferMs = _mainBufEffectiveMs; s.AuxBufferMs = _auxBufEffectiveMs;

                s.MainDefaultPeriodMs = _defMainMs; s.MainMinimumPeriodMs = _minMainMs;
                s.AuxDefaultPeriodMs  = _defAuxMs;  s.AuxMinimumPeriodMs  = _minAuxMs;

                // 直通标记与描述
                s.MainPassthrough = (_resMain == null);
                s.AuxPassthrough  = (_resAux  == null);

                s.MainPassDesc = _mainPath == PathType.PassthroughExclusive ? "独占直通"
                                : _mainPath == PathType.PassthroughSharedMix ? "共享混音直通"
                                : "重采样";

                s.AuxPassDesc = _auxPath == PathType.PassthroughExclusive ? "独占直通"
                               : _auxPath == PathType.PassthroughSharedMix ? "共享混音直通"
                               : "重采样";

                s.AuxQuality = _cfg != null ? _cfg.AuxResamplerQuality : 40;
            }
            catch { }
            return s;
        }

        static string Fmt(WaveFormat f)
        {
            if (f == null) return "-";
            string enc = f.Encoding == WaveFormatEncoding.IeeeFloat ? "32f" : (f.BitsPerSample + "bit");
            return f.SampleRate + " Hz / " + enc + " / " + f.Channels + "ch";
        }

        void EnsureAutoStart(bool on)
        {
            try
            {
                using (var rk = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
                {
                    if (on) rk.SetValue("MirrorAudio", Application.ExecutablePath);
                    else rk.DeleteValue("MirrorAudio", false);
                }
            } catch { }
        }

        // IMMNotificationClient（热插拔）
        public void OnDeviceStateChanged(string deviceId, DeviceState newState) { DebouncedRestart(); }
        public void OnDeviceAdded(string pwstrDeviceId) { DebouncedRestart(); }
        public void OnDeviceRemoved(string deviceId) { DebouncedRestart(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId) { DebouncedRestart(); }
        public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key) { }
        void DebouncedRestart() { _debounce.Stop(); _debounce.Start(); }
    }

    // ============== 配置存取（极简 JSON-free，本地 ini 样式） ==============

    static class ConfigStore
    {
        static readonly string PathCfg = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MirrorAudio.cfg");

        public static AppSettings Load()
        {
            var c = new AppSettings();
            try
            {
                if (!File.Exists(PathCfg)) return c;
                foreach (var line in File.ReadAllLines(PathCfg, Encoding.UTF8))
                {
                    var s = line.Trim(); if (s.Length == 0 || s.StartsWith("#")) continue;
                    var i = s.IndexOf('='); if (i <= 0) continue;
                    var k = s.Substring(0, i).Trim(); var v = s.Substring(i + 1).Trim();
                    if (k == "InputDeviceId") c.InputDeviceId = v;
                    else if (k == "MainDeviceId") c.MainDeviceId = v;
                    else if (k == "AuxDeviceId") c.AuxDeviceId = v;
                    else if (k == "MainShare") c.MainShare = (ShareModeOption)Enum.Parse(typeof(ShareModeOption), v);
                    else if (k == "AuxShare") c.AuxShare = (ShareModeOption)Enum.Parse(typeof(ShareModeOption), v);
                    else if (k == "MainSync") c.MainSync = (SyncModeOption)Enum.Parse(typeof(SyncModeOption), v);
                    else if (k == "AuxSync") c.AuxSync = (SyncModeOption)Enum.Parse(typeof(SyncModeOption), v);
                    else if (k == "MainRate") c.MainRate = int.Parse(v);
                    else if (k == "MainBits") c.MainBits = int.Parse(v);
                    else if (k == "MainBufMs") c.MainBufMs = int.Parse(v);
                    else if (k == "AuxRate") c.AuxRate = int.Parse(v);
                    else if (k == "AuxBits") c.AuxBits = int.Parse(v);
                    else if (k == "AuxBufMs") c.AuxBufMs = int.Parse(v);
                    else if (k == "AutoStart") c.AutoStart = v == "true";
                    else if (k == "EnableLogging") c.EnableLogging = v == "true";
                    else if (k == "AuxResamplerQuality") c.AuxResamplerQuality = int.Parse(v);
                }
            } catch { }
            return c;
        }

        public static void Save(AppSettings c)
        {
            try
            {
                var sb = new StringBuilder();
                Action<string, object> w = (k, v) => sb.AppendLine(k + "=" + v);
                w("InputDeviceId", c.InputDeviceId ?? "");
                w("MainDeviceId", c.MainDeviceId ?? "");
                w("AuxDeviceId", c.AuxDeviceId ?? "");
                w("MainShare", c.MainShare);
                w("AuxShare", c.AuxShare);
                w("MainSync", c.MainSync);
                w("AuxSync", c.AuxSync);
                w("MainRate", c.MainRate);
                w("MainBits", c.MainBits);
                w("MainBufMs", c.MainBufMs);
                w("AuxRate", c.AuxRate);
                w("AuxBits", c.AuxBits);
                w("AuxBufMs", c.AuxBufMs);
                w("AutoStart", c.AutoStart ? "true" : "false");
                w("EnableLogging", c.EnableLogging ? "true" : "false");
                w("AuxResamplerQuality", c.AuxResamplerQuality);
                File.WriteAllText(PathCfg, sb.ToString(), Encoding.UTF8);
            } catch { }
        }

        public static bool Equals(AppSettings a, AppSettings b)
        {
            if (a == null || b == null) return false;
            return a.InputDeviceId == b.InputDeviceId &&
                   a.MainDeviceId == b.MainDeviceId &&
                   a.AuxDeviceId == b.AuxDeviceId &&
                   a.MainShare == b.MainShare &&
                   a.AuxShare == b.AuxShare &&
                   a.MainSync == b.MainSync &&
                   a.AuxSync == b.AuxSync &&
                   a.MainRate == b.MainRate &&
                   a.MainBits == b.MainBits &&
                   a.MainBufMs == b.MainBufMs &&
                   a.AuxRate == b.AuxRate &&
                   a.AuxBits == b.AuxBits &&
                   a.AuxBufMs == b.AuxBufMs &&
                   a.AutoStart == b.AutoStart &&
                   a.EnableLogging == b.EnableLogging &&
                   a.AuxResamplerQuality == b.AuxResamplerQuality;
        }
    }

    // ============== 轻量日志 ==============

    static class Logger
    {
        static string _path;
        public static bool Enabled;

        public static void Init()
        {
            try
            {
                _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MirrorAudio.log");
                File.WriteAllText(_path, DateTime.Now + " start\n");
            } catch { }
        }

        public static void Log(string s)
        {
            if (!Enabled) return;
            try { File.AppendAllText(_path, DateTime.Now.ToString("HH:mm:ss.fff ") + s + "\n"); } catch { }
        }
    }
}
