using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try { Logger.Init(); } catch { }
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Log("unhandled: " + e.ExceptionObject);
            Application.ThreadException += (s, e) => Logger.Log("ui-ex: " + e.Exception);

            Application.Run(new TrayContext()); // 纯托盘，无窗体白屏
        }
    }

    // ---------- 配置 / 状态 ----------
    public enum ShareModeOption { Auto, Exclusive, Shared }
    public enum SyncModeOption  { Auto, Event, Polling }
    public enum PathType
    {
        None,
        PassthroughExclusive,
        PassthroughSharedMix,
        ResampledExclusive,
        ResampledShared
    }

    public sealed class AppSettings
    {
        public string InputDeviceId, MainDeviceId, AuxDeviceId;
        public ShareModeOption MainShare = ShareModeOption.Auto, AuxShare = ShareModeOption.Auto;
        public SyncModeOption  MainSync  = SyncModeOption.Auto,  AuxSync  = SyncModeOption.Auto;
        public int MainRate = 192000, MainBits = 24, MainBufMs = 12;
        public int AuxRate  =  48000, AuxBits  = 16, AuxBufMs  = 160;
        public bool AutoStart = false, EnableLogging = false;
        public int AuxResamplerQuality = 40; // 30/40/50
    }

    public sealed class StatusSnapshot
    {
        public bool Running;
        public string InputDevice, InputRole, InputFormat;
        public string MainDevice, MainMode, MainSync, MainFormat;
        public string AuxDevice,  AuxMode,  AuxSync,  AuxFormat;
        public int MainBufferMs, AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs, AuxDefaultPeriodMs, AuxMinimumPeriodMs;
        public bool MainPassthrough, AuxPassthrough;
        public string MainPassDesc, AuxPassDesc;
        public int AuxQuality;
    }

    // ---------- 托盘上下文 ----------
    public sealed class TrayContext : ApplicationContext, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly System.Windows.Forms.Timer _debounce = new System.Windows.Forms.Timer { Interval = 600 };
        readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        AppSettings _cfg = ConfigStore.Load();
        readonly Dictionary<string, Tuple<double,double>> _periodCache = new Dictionary<string, Tuple<double,double>>(8);
        static readonly int[] CommonRates = new[] { 48000, 96000, 192000, 44100 };

        // 输入
        WasapiCapture _cap;
        BufferedWaveProvider _bufMain, _bufAux;

        // 输出-主
        MMDevice _outMain; WasapiOut _mainOut; MediaFoundationResampler _resMain;
        bool _mainIsExclusive, _mainEventSyncUsed; int _mainBufEffectiveMs; PathType _mainPath = PathType.None;

        // 输出-副
        MMDevice _outAux;  WasapiOut _auxOut;  MediaFoundationResampler _resAux;
        bool _auxIsExclusive,  _auxEventSyncUsed;  int _auxBufEffectiveMs;  PathType _auxPath = PathType.None;

        // 状态文本
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-";
        string _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs = 10, _minMainMs = 2, _defAuxMs = 10, _minAuxMs = 2;

        public TrayContext()
        {
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            try
            {
                var ico = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                _tray.Icon = File.Exists(ico) ? new Icon(ico) : Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            } catch { _tray.Icon = SystemIcons.Application; }
            _tray.Visible = true; _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s,e)=>StartOrRestart());
            var miStop  = new ToolStripMenuItem("停止(&T)",    null, (s,e)=>Stop());
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, (s,e)=>OnSettings());
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s,e)=>Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)",    null, (s,e)=>{ Stop(); _tray.Visible=false; _tray.Dispose(); ExitThread(); });
            _menu.Items.AddRange(new ToolStripItem[]{ miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += (s,e)=>{ _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            Logger.Enabled = _cfg.EnableLogging;
            StartOrRestart();
        }

        void OnSettings()
        {
            using (var f=new SettingsForm(_cfg, GetStatusSnapshot))
            {
                if (f.ShowDialog()==DialogResult.OK && f.Result!=null)
                {
                    var old=_cfg; _cfg=f.Result;
                    Logger.Enabled=_cfg.EnableLogging;
                    EnsureAutoStart(_cfg.AutoStart);
                    if (!ConfigStore.Equals(old,_cfg)) ConfigStore.Save(_cfg);
                    _debounce.Stop(); _debounce.Start();
                }
            }
        }

        // ---------- 启停 ----------
        void StartOrRestart()
        {
            Stop();
            try
            {
                // 输入设备
                var capDev = PickInput(_cfg.InputDeviceId);
                if (capDev==null) { MessageBox.Show("未找到输入设备。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information); return; }
                _inDevName = capDev.FriendlyName;
                bool isLoopback = capDev.DataFlow==DataFlow.Render;

                _cap = isLoopback ? (WasapiCapture) new WasapiLoopbackCapture(capDev) : new WasapiCapture(capDev);
                _cap.ShareMode = AudioClientShareMode.Shared;
                _cap.DataAvailable += OnData;
                _cap.RecordingStopped += (s,e)=> Logger.Log("capture stopped: "+e.Exception);
                var inFmt = _cap.WaveFormat;
                _inRoleStr = isLoopback? "环回":"录音";
                _inFmtStr  = Fmt(inFmt);

                // 输出设备
                _outMain = PickRender(_cfg.MainDeviceId);
                _outAux  = PickRender(_cfg.AuxDeviceId);
                if (_outMain==null || _outAux==null) { MessageBox.Show("请在设置中选择主/副输出设备。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information); return; }

                var perM = GetOrCachePeriod(_outMain); _defMainMs = perM.Item1; _minMainMs = perM.Item2;
                var perA = GetOrCachePeriod(_outAux ); _defAuxMs  = perA.Item1; _minAuxMs  = perA.Item2;

                // 环形缓存（紧凑，不改延迟）
                _bufMain = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true,
                    BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs*6, 96)) };
                _bufAux  = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true,
                    BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs*4, 120)) };

                // 主/副：统一初始化（独占直通优先；失败回退共享直通→共享重采样）
                bool triedMainEx=false, triedAuxEx=false;

                InitPath(
                    which:"主通道",
                    cfgShare:_cfg.MainShare, cfgSync:_cfg.MainSync,
                    userRate:_cfg.MainRate, userBits:_cfg.MainBits,
                    wantMs:_cfg.MainBufMs, defMs:_defMainMs, minMs:_minMainMs,
                    dev:_outMain, inFmt:inFmt, buf:_bufMain,
                    resRef: ref _resMain, outRef: ref _mainOut,
                    isExclusive: out _mainIsExclusive, usedEvent: out _mainEventSyncUsed,
                    effMs: out _mainBufEffectiveMs, fmtStr: out _mainFmtStr, path: out _mainPath,
                    auxResamplerQuality: 50,
                    preferExclusiveDefault:true,
                    exclusiveAttempted: out triedMainEx
                );

                InitPath(
                    which:"副通道",
                    cfgShare:_cfg.AuxShare, cfgSync:_cfg.AuxSync,
                    userRate:_cfg.AuxRate, userBits:_cfg.AuxBits,
                    wantMs:_cfg.AuxBufMs, defMs:_defAuxMs, minMs:_minAuxMs,
                    dev:_outAux, inFmt:inFmt, buf:_bufAux,
                    resRef: ref _resAux, outRef: ref _auxOut,
                    isExclusive: out _auxIsExclusive, usedEvent: out _auxEventSyncUsed,
                    effMs: out _auxBufEffectiveMs, fmtStr: out _auxFmtStr, path: out _auxPath,
                    auxResamplerQuality: Clamp(_cfg.AuxResamplerQuality,30,50),
                    preferExclusiveDefault:true, // Auto 也先试独占
                    exclusiveAttempted: out triedAuxEx
                );

                // 起播
                _cap.StartRecording();
                if (_mainOut!=null) _mainOut.Play();
                if (_auxOut !=null) _auxOut.Play();

                // 回退提示（仅当尝试过独占但失败）
                MaybeNotifyFallback("主通道", _cfg.MainShare, triedMainEx, _mainIsExclusive, _mainPath);
                MaybeNotifyFallback("副通道", _cfg.AuxShare,  triedAuxEx,  _auxIsExclusive,  _auxPath);
            }
            catch (Exception ex)
            {
                Logger.Log("Start failed: "+ex);
                MessageBox.Show("启动失败："+ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        void Stop()
        {
            try { _mainOut?.Stop(); } catch { }
            try { _auxOut ?.Stop(); } catch { }
            try { _cap    ?.StopRecording(); } catch { }

            DisposeSafe(ref _resMain); DisposeSafe(ref _resAux);
            DisposeSafe(ref _mainOut); DisposeSafe(ref _auxOut);
            DisposeSafe(ref _cap);
            _mainPath = PathType.None; _auxPath = PathType.None;
        }

        // ---------- 合并后的通用初始化 ----------
        void InitPath(
            string which,
            ShareModeOption cfgShare, SyncModeOption cfgSync,
            int userRate, int userBits,
            int wantMs, double defMs, double minMs,
            MMDevice dev, WaveFormat inFmt, BufferedWaveProvider buf,
            ref MediaFoundationResampler resRef, ref WasapiOut outRef,
            out bool isExclusive, out bool usedEvent, out int effMs, out string fmtStr, out PathType path,
            int auxResamplerQuality, bool preferExclusiveDefault, out bool exclusiveAttempted)
        {
            isExclusive=false; usedEvent=false; effMs=0; fmtStr="-"; path=PathType.None; exclusiveAttempted=false;
            if (dev==null || inFmt==null || buf==null) return;

            // 只要不是“明确选共享”，Auto/Exclusive 都先尝试独占
            bool wantExclusive = (cfgShare != ShareModeOption.Shared) &&
                                 ((cfgShare==ShareModeOption.Exclusive) || (cfgShare==ShareModeOption.Auto && preferExclusiveDefault));

            var candidates = BuildExclusiveCandidates(inFmt, userRate, userBits);

            if (wantExclusive)
            {
                exclusiveAttempted = true;
                for (int i=0;i<candidates.Count;i++)
                {
                    var fmt = candidates[i];
                    bool ok = SupportsExclusive(dev, fmt);
                    if (!ok)
                    {
                        if (Logger.Enabled) Logger.Log(which + " exclusive reject: " + Fmt(fmt));
                        continue;
                    }

                    int ms = Buf(wantMs, true, defMs, minMs);
                    IWaveProvider src = buf;
                    if (!Eq(buf.WaveFormat, fmt))
                    {
                        resRef = new MediaFoundationResampler(buf, fmt){ ResamplerQuality = auxResamplerQuality };
                        src = resRef;
                    }

                    var wo = CreateOut(dev, AudioClientShareMode.Exclusive, cfgSync, ms, src, out usedEvent, fmt);
                    if (wo != null)
                    {
                        outRef = wo; isExclusive=true; effMs=ms; fmtStr=Fmt(fmt);
                        path = Eq(buf.WaveFormat, fmt) ? PathType.PassthroughExclusive : PathType.ResampledExclusive;
                        if (Logger.Enabled) Logger.Log(which + " exclusive OK: " + Fmt(fmt) + " , " + (usedEvent?"event":"polling") + ", " + ms + "ms");
                        return;
                    }
                    DisposeSafe(ref resRef);
                    if (Logger.Enabled) Logger.Log(which + " exclusive init-fail: " + Fmt(fmt));
                }
            }

            // 共享回退：混音直通优先，否则重采样
            int msS = Buf(wantMs, false, defMs);
            WaveFormat mix=null; try{ mix=dev.AudioClient.MixFormat; } catch { }
            IWaveProvider shareSrc = buf;

            if (mix != null && !Eq(mix, buf.WaveFormat))
            {
                resRef = new MediaFoundationResampler(buf, mix){ ResamplerQuality = auxResamplerQuality };
                shareSrc = resRef; path = PathType.ResampledShared;
            }
            else
            {
                path = PathType.PassthroughSharedMix;
            }

            var woS = CreateOut(dev, AudioClientShareMode.Shared, cfgSync, msS, shareSrc, out usedEvent, null);
            if (woS != null)
            {
                outRef = woS; effMs=msS; fmtStr=Fmt(mix??buf.WaveFormat);
                if (Logger.Enabled) Logger.Log(which + " shared " + (path==PathType.PassthroughSharedMix?"passthrough":"resample") + " , " + (usedEvent?"event":"polling") + ", " + msS + "ms");
            }
        }

        List<WaveFormat> BuildExclusiveCandidates(WaveFormat inFmt, int userRate, int userBits)
        {
            var list = new List<WaveFormat>(12);
            // 1) 输入直通
            list.Add(inFmt);

            // 2) 同采样率三件套：32f、24-bit packed、32-bit PCM（很多 SPDIF/HDMI 用 32 容器承载 24 有效位）
            if (inFmt.SampleRate>0 && inFmt.Channels>0)
            {
                var f32f   = WaveFormat.CreateIeeeFloatWaveFormat(inFmt.SampleRate, inFmt.Channels);
                var f24    = Pcm24(inFmt.SampleRate, inFmt.Channels);
                var f32pcm = Pcm32(inFmt.SampleRate, inFmt.Channels);
                if (!Eq(f32f,   inFmt)) list.Add(f32f);
                if (!Eq(f24,    inFmt)) list.Add(f24);
                if (!Eq(f32pcm, inFmt)) list.Add(f32pcm);
            }

            // 3) 用户设定（仅独占生效）
            var user = WaveFormatFromUser(userRate, userBits, inFmt.Channels, false);
            if (user != null && !list.Any(f=>Eq(f,user))) list.Add(user);

            // 4) 常见采样率（每个三件套）。48/96 优先，再 192/44.1
            for (int i=0;i<CommonRates.Length;i++)
            {
                int r = CommonRates[i]; if (r==inFmt.SampleRate) continue;
                var f24    = Pcm24(r, inFmt.Channels);
                var f32pcm = Pcm32(r, inFmt.Channels);
                var f32f   = WaveFormat.CreateIeeeFloatWaveFormat(r, inFmt.Channels);
                if (!list.Any(f=>Eq(f,f24)))    list.Add(f24);
                if (!list.Any(f=>Eq(f,f32pcm))) list.Add(f32pcm);
                if (!list.Any(f=>Eq(f,f32f)))   list.Add(f32f);
            }
            return list;
        }

        static WaveFormat Pcm24(int rate, int ch) { return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, ch, rate*ch*3, ch*3, 24); }
        static WaveFormat Pcm32(int rate, int ch) { return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, ch, rate*ch*4, ch*4, 32); }

        // ---------- 数据回调 ----------
        void OnData(object s, WaveInEventArgs e)
        {
            try { _bufMain?.AddSamples(e.Buffer,0,e.BytesRecorded); _bufAux?.AddSamples(e.Buffer,0,e.BytesRecorded); }
            catch (Exception ex) { Logger.Log("OnData: "+ex.Message); }
        }

        // ---------- 工具 ----------
        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption sync, int ms, IWaveProvider src, out bool eventUsed, WaveFormat forceFormat)
        {
            eventUsed=false;
            try
            {
                bool useEvent = (sync==SyncModeOption.Event) || (sync==SyncModeOption.Auto && PreferEvent());
                var wo = new WasapiOut(dev, mode, useEvent, ms);
                eventUsed = useEvent;

                if (forceFormat!=null && !Eq(src.WaveFormat, forceFormat))
                    src = new MediaFoundationResampler(src, forceFormat){ ResamplerQuality = 50 };

                wo.Init(src);
                return wo;
            }
            catch (Exception ex) { Logger.Log("CreateOut: "+ex.Message); return null; }
        }

        static bool PreferEvent() { return true; }

        static int Buf(int wantMs, bool isExclusive, double defMs, double minMs = 2)
        {
            int ms = Math.Max(4, wantMs);
            if (isExclusive)
            {
                int m = Math.Max(1, (int)Math.Round(minMs));
                int k = (int)Math.Ceiling(ms / (double)m);
                ms = Math.Max(k*m, (int)Math.Ceiling(minMs));
            }
            return ms;
        }

        static int Clamp(int v,int lo,int hi){ return v<lo?lo:(v>hi?hi:v); }

        WaveFormat WaveFormatFromUser(int rate,int bits,int ch,bool preferFloat)
        {
            try
            {
                if (preferFloat && bits>=32) return WaveFormat.CreateIeeeFloatWaveFormat(rate,ch);
                if (bits==24) return Pcm24(rate,ch);
                if (bits==32 && preferFloat) return WaveFormat.CreateIeeeFloatWaveFormat(rate,ch);
                if (bits==32) return Pcm32(rate,ch);
                return new WaveFormat(rate,bits,ch);
            } catch { return null; }
        }

        static bool Eq(WaveFormat a,WaveFormat b)
        {
            if (a==null || b==null) return false;
            return a.SampleRate==b.SampleRate && a.BitsPerSample==b.BitsPerSample &&
                   a.Channels==b.Channels && a.Encoding==b.Encoding;
        }

        Tuple<double,double> GetOrCachePeriod(MMDevice dev)
        {
            if (dev==null) return Tuple.Create(10.0,2.0);
            Tuple<double,double> t;
            if (_periodCache.TryGetValue(dev.ID, out t)) return t;
            double defMs=10, minMs=2;
            try {
                var ac=dev.AudioClient;
                long def100ns=ac.DefaultDevicePeriod, min100ns=ac.MinimumDevicePeriod;
                defMs = def100ns/10000.0; minMs = min100ns/10000.0;
            } catch { }
            t = Tuple.Create(defMs,minMs); _periodCache[dev.ID]=t; return t;
        }

        MMDevice PickInput(string id)
        {
            try{
                if(!string.IsNullOrEmpty(id)){
                    var d=_mm.GetDevice(id);
                    if (d.State==DeviceState.Active && (d.DataFlow==DataFlow.Capture || d.DataFlow==DataFlow.Render)) return d;
                }
            } catch{}
            try{ return _mm.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console); }catch{}
            try{ return _mm.GetDefaultAudioEndpoint(DataFlow.Render , Role.Console); }catch{}
            return null;
        }

        MMDevice PickRender(string id)
        {
            try{
                if(!string.IsNullOrEmpty(id)){
                    var d=_mm.GetDevice(id);
                    if (d.State==DeviceState.Active && d.DataFlow==DataFlow.Render) return d;
                }
            } catch{}
            try{ return _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia); }catch{}
            return null;
        }

        bool SupportsExclusive(MMDevice dev, WaveFormat fmt)
        {
            try { return dev.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, fmt); }
            catch { return false; }
        }

        static void DisposeSafe<T>(ref T o) where T:class,IDisposable
        { try{ o?.Dispose(); } catch{} finally{ o=null; } }

        // ---------- 状态提供 ----------
        public StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot();
            try
            {
                s.Running = _cap!=null && _mainOut!=null && _auxOut!=null;
                s.InputDevice=_inDevName; s.InputRole=_inRoleStr; s.InputFormat=_inFmtStr;

                s.MainDevice=_outMain?.FriendlyName; s.AuxDevice=_outAux?.FriendlyName;
                s.MainMode = _mainIsExclusive? "独占":"共享";
                s.AuxMode  = _auxIsExclusive ? "独占":"共享";
                s.MainSync = _mainEventSyncUsed? "事件":"轮询";
                s.AuxSync  = _auxEventSyncUsed ? "事件":"轮询";

                s.MainFormat=_mainFmtStr; s.AuxFormat=_auxFmtStr;
                s.MainBufferMs=_mainBufEffectiveMs; s.AuxBufferMs=_auxBufEffectiveMs;

                s.MainDefaultPeriodMs=_defMainMs; s.MainMinimumPeriodMs=_minMainMs;
                s.AuxDefaultPeriodMs=_defAuxMs;  s.AuxMinimumPeriodMs=_minAuxMs;

                s.MainPassthrough = (_resMain==null);
                s.AuxPassthrough  = (_resAux ==null);

                s.MainPassDesc = _mainPath==PathType.PassthroughExclusive ? "独占直通"
                              : _mainPath==PathType.PassthroughSharedMix ? "共享直通"
                              : _mainPath==PathType.ResampledExclusive   ? "独占重采样"
                              : _mainPath==PathType.ResampledShared      ? "共享重采样":"-";

                s.AuxPassDesc  = _auxPath==PathType.PassthroughExclusive ? "独占直通"
                              : _auxPath==PathType.PassthroughSharedMix ? "共享直通"
                              : _auxPath==PathType.ResampledExclusive   ? "独占重采样"
                              : _auxPath==PathType.ResampledShared      ? "共享重采样":"-";

                s.AuxQuality = _cfg!=null ? _cfg.AuxResamplerQuality : 40;
            } catch { }
            return s;
        }

        // ---------- 回退提示 ----------
        void MaybeNotifyFallback(string which, ShareModeOption cfgShare, bool exclusiveAttempted, bool gotExclusive, PathType path)
        {
            // 只在尝试过独占但最终没拿到时提示
            if (!exclusiveAttempted || gotExclusive) return;

            string desc = (path==PathType.PassthroughSharedMix) ? "共享直通"
                       : (path==PathType.ResampledShared)       ? "共享重采样"
                       : "共享";
            TrayTip($"{which}未能独占，已回退为「{desc}」。可尝试同采样率的 32-bit PCM 或 32f。");
            Logger.Log($"{which} exclusive fallback -> {desc}");
        }

        void TrayTip(string text)
        {
            try{ _tray.BalloonTipTitle="MirrorAudio"; _tray.BalloonTipText=text; _tray.ShowBalloonTip(2500); } catch{}
        }

        // ==== IMMNotificationClient：设备热插拔事件（事件驱动自愈）====
        public void OnDeviceStateChanged(string deviceId, DeviceState newState) { DebouncedRestart(); }
        public void OnDeviceAdded(string pwstrDeviceId) { DebouncedRestart(); }
        public void OnDeviceRemoved(string deviceId) { DebouncedRestart(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId) { DebouncedRestart(); }
        public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key) { }

        // 统一的去抖重启
        void DebouncedRestart(){ _debounce.Stop(); _debounce.Start(); }

        // 自启动
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

        // 格式字符串
        static string Fmt(WaveFormat f)
        {
            if (f == null) return "-";
            string enc = (f.Encoding == WaveFormatEncoding.IeeeFloat) ? "32f" : (f.BitsPerSample + "bit");
            return f.SampleRate + " Hz / " + enc + " / " + f.Channels + "ch";
        }
    }

    // ---------- 配置存取 ----------
    static class ConfigStore
    {
        static readonly string PathCfg = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MirrorAudio.cfg");

        public static AppSettings Load()
        {
            var c = new AppSettings();
            try{
                if(!File.Exists(PathCfg)) return c;
                var lines = File.ReadAllLines(PathCfg, Encoding.UTF8);
                for(int i=0;i<lines.Length;i++)
                {
                    var s=lines[i].Trim(); if (s.Length==0 || s.StartsWith("#")) continue;
                    int p=s.IndexOf('='); if (p<=0) continue;
                    var k=s.Substring(0,p).Trim(); var v=s.Substring(p+1).Trim();
                    if (k=="InputDeviceId") c.InputDeviceId=v;
                    else if (k=="MainDeviceId") c.MainDeviceId=v;
                    else if (k=="AuxDeviceId") c.AuxDeviceId=v;
                    else if (k=="MainShare") c.MainShare=(ShareModeOption)Enum.Parse(typeof(ShareModeOption),v);
                    else if (k=="AuxShare") c.AuxShare=(ShareModeOption)Enum.Parse(typeof(ShareModeOption),v);
                    else if (k=="MainSync") c.MainSync=(SyncModeOption)Enum.Parse(typeof(SyncModeOption),v);
                    else if (k=="AuxSync") c.AuxSync=(SyncModeOption)Enum.Parse(typeof(SyncModeOption),v);
                    else if (k=="MainRate") c.MainRate=int.Parse(v);
                    else if (k=="MainBits") c.MainBits=int.Parse(v);
                    else if (k=="MainBufMs") c.MainBufMs=int.Parse(v);
                    else if (k=="AuxRate") c.AuxRate=int.Parse(v);
                    else if (k=="AuxBits") c.AuxBits=int.Parse(v);
                    else if (k=="AuxBufMs") c.AuxBufMs=int.Parse(v);
                    else if (k=="AutoStart") c.AutoStart = (v=="true");
                    else if (k=="EnableLogging") c.EnableLogging = (v=="true");
                    else if (k=="AuxResamplerQuality") c.AuxResamplerQuality=int.Parse(v);
                }
            } catch {}
            return c;
        }

        public static void Save(AppSettings c)
        {
            try{
                var sb=new StringBuilder(); Action<string,object> W=(k,v)=>sb.AppendLine(k+"="+v);
                W("InputDeviceId", c.InputDeviceId??"");
                W("MainDeviceId",  c.MainDeviceId ??"");
                W("AuxDeviceId",   c.AuxDeviceId  ??"");
                W("MainShare", c.MainShare);  W("AuxShare", c.AuxShare);
                W("MainSync",  c.MainSync );  W("AuxSync",  c.AuxSync );
                W("MainRate",  c.MainRate );  W("MainBits", c.MainBits);
                W("MainBufMs", c.MainBufMs);
                W("AuxRate",   c.AuxRate  );  W("AuxBits",  c.AuxBits );
                W("AuxBufMs",  c.AuxBufMs );
                W("AutoStart", c.AutoStart? "true":"false");
                W("EnableLogging", c.EnableLogging? "true":"false");
                W("AuxResamplerQuality", c.AuxResamplerQuality);
                File.WriteAllText(PathCfg, sb.ToString(), Encoding.UTF8);
            } catch {}
        }

        public static bool Equals(AppSettings a, AppSettings b)
        {
            if (a==null || b==null) return false;
            return a.InputDeviceId==b.InputDeviceId &&
                   a.MainDeviceId ==b.MainDeviceId  &&
                   a.AuxDeviceId  ==b.AuxDeviceId   &&
                   a.MainShare==b.MainShare && a.AuxShare==b.AuxShare &&
                   a.MainSync ==b.MainSync  && a.AuxSync ==b.AuxSync  &&
                   a.MainRate ==b.MainRate  && a.MainBits==b.MainBits && a.MainBufMs==b.MainBufMs &&
                   a.AuxRate  ==b.AuxRate   && a.AuxBits ==b.AuxBits  && a.AuxBufMs ==b.AuxBufMs  &&
                   a.AutoStart==b.AutoStart && a.EnableLogging==b.EnableLogging &&
                   a.AuxResamplerQuality==b.AuxResamplerQuality;
        }
    }

    // ---------- 轻量日志 ----------
    static class Logger
    {
        static string _path; public static bool Enabled;
        public static void Init(){ try{ _path=Path.Combine(Path.GetTempPath(),"MirrorAudio.log"); File.WriteAllText(_path, DateTime.Now+" start\n"); } catch{} }
        public static void Log(string s){ if(!Enabled) return; try{ File.AppendAllText(_path, DateTime.Now.ToString("HH:mm:ss.fff ")+s+"\n"); } catch{} }
    }
}
