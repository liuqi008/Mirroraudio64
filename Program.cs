
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
    // —— 日志 —— //
    static class Logger
    {
        public static bool Enabled; // 默认关闭
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        static readonly string CrashPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
        public static void Info(string s){ if(!Enabled) return; try{ File.AppendAllText(LogPath,"["+DateTime.Now.ToString("HH:mm:ss")+"] "+s+"\r\n"); }catch{} }
        public static void Crash(string where, Exception ex){ if(ex==null) return; try{ File.AppendAllText(CrashPath,$"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n"); }catch{} }
    }

    // —— 入口 —— //
    static class Program
    {
        static Mutex _mtx;
        [STAThread]
        static void Main()
        {
            bool ok; _mtx=new Mutex(true,"Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}",out ok); if(!ok) return;
            Application.ThreadException+=(s,e)=>Logger.Crash("UI",e.Exception);
            AppDomain.CurrentDomain.UnhandledException+=(s,e)=>Logger.Crash("NonUI",e.ExceptionObject as Exception);
            try{ MediaFoundationApi.Startup(); }catch{}
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            using(var app=new TrayApp()) Application.Run();
            try{ MediaFoundationApi.Shutdown(); }catch{}
        }
    }

    // —— 配置枚举 —— //
    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    // 输入环回申请策略（B 方案）
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

    // —— 配置 —— //
    [DataContract]
    public sealed class AppSettings
    {
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare=ShareModeOption.Auto, AuxShare=ShareModeOption.Shared;
        [DataMember] public SyncModeOption  MainSync =SyncModeOption.Auto,  AuxSync =SyncModeOption.Auto;
        [DataMember] public int MainRate=192000, MainBits=24, MainBufMs=12;
        [DataMember] public int AuxRate =48000,  AuxBits =16,  AuxBufMs =150;
        [DataMember] public bool AutoStart=false, EnableLogging=false;

        // —— 新增：输入环回格式策略（B 方案） —— //
        [DataMember] public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        [DataMember] public int InputCustomSampleRate = 96000; // 当 Strategy=Custom 时生效
        [DataMember] public int InputCustomBitDepth  = 24;     // 16/24/32(=float)
    }

    // —— 状态 —— //
    public sealed class StatusSnapshot
    {
        public bool Running;
        public int MainBufferRequestedMs, AuxBufferRequestedMs;
        public double MainAlignedMultiple, AuxAlignedMultiple;
        public double MainBufferMultiple, AuxBufferMultiple;
        public string InputRole,InputFormat,InputDevice;
        public string InputRequested,InputAccepted,InputMix;
        public string MainDevice,AuxDevice,MainMode,AuxMode,MainSync,AuxSync,MainFormat,AuxFormat;
        public int MainBufferMs,AuxBufferMs;
        public double MainDefaultPeriodMs,MainMinimumPeriodMs,AuxDefaultPeriodMs,AuxMinimumPeriodMs;
        public bool MainNoSRC,AuxNoSRC,MainResampling,AuxResampling;
    }

    static class Config
    {
        static readonly string Dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"MirrorAudio");
        static readonly string FilePath=Path.Combine(Dir,"settings.json");
        public static AppSettings Load(){ try{ if(!File.Exists(FilePath)) return new AppSettings(); using(var fs=File.OpenRead(FilePath)) return (AppSettings)new DataContractJsonSerializer(typeof(AppSettings)).ReadObject(fs);}catch{ return new AppSettings(); }}
        public static void Save(AppSettings s){ try{ if(!Directory.Exists(Dir)) Directory.CreateDirectory(Dir); using(var fs=File.Create(FilePath)) new DataContractJsonSerializer(typeof(AppSettings)).WriteObject(fs,s);}catch{} }
    }

    // —— 主体 —— //
    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray=new NotifyIcon();
        readonly ContextMenuStrip _menu=new ContextMenuStrip();
        readonly WinFormsTimer _debounce=new WinFormsTimer{ Interval=400 };

        AppSettings _cfg=Config.Load();
        MMDeviceEnumerator _mm=new MMDeviceEnumerator();

        MMDevice _inDev,_outMain,_outAux;
        IWaveIn _capture; BufferedWaveProvider _bufMain,_bufAux;
        IWaveProvider _srcMain,_srcAux; WasapiOut _mainOut,_auxOut;
        MediaFoundationResampler _resMain,_resAux;

        bool _running,_mainIsExclusive,_mainEventSyncUsed,_auxIsExclusive,_auxEventSyncUsed;
        int _mainBufEffectiveMs,_auxBufEffectiveMs;
        string _inRoleStr="-",_inFmtStr="-",_inDevName="-",_mainFmtStr="-",_auxFmtStr="-";
        string _inReqStr="-", _inAccStr="-", _inMixStr="-";
        bool _mainNoSRC, _auxNoSRC, _mainResampling, _auxResampling;
        double _defMainMs=10,_minMainMs=2,_defAuxMs=10,_minAuxMs=2;

        readonly Dictionary<string,Tuple<double,double>> _periodCache=new Dictionary<string,Tuple<double,double>>(4);

        public TrayApp()
        {
            Logger.Enabled=_cfg.EnableLogging;
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            // 托盘图标：优先用 Assets\MirrorAudio.ico，其次用 exe 关联，最后系统默认
            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                if (File.Exists(icoPath))
                    _tray.Icon = new Icon(icoPath);
                else
                    _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch
            {
                _tray.Icon = SystemIcons.Application;
            }

            _tray.Visible = true;
            _tray.Text = "MirrorAudio";

            var miStart=new ToolStripMenuItem("启动/重启(&S)",null,(s,e)=>StartOrRestart());
            var miStop =new ToolStripMenuItem("停止(&T)",   null,(s,e)=>Stop());
            var miSet  =new ToolStripMenuItem("设置(&G)...",null,(s,e)=>OnSettings());
            var miLog  =new ToolStripMenuItem("打开日志目录",null,(s,e)=>Process.Start("explorer.exe",Path.GetTempPath()));
            var miExit =new ToolStripMenuItem("退出(&X)",   null,(s,e)=>{ Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[]{ miStart,miStop,new ToolStripSeparator(),miSet,miLog,new ToolStripSeparator(),miExit });
            _tray.ContextMenuStrip=_menu;

            _debounce.Tick+=(s,e)=>{ _debounce.Stop(); StartOrRestart(); };

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        void OnSettings()
        {
            using(var f=new SettingsForm(_cfg,GetStatusSnapshot))
            {
                if(f.ShowDialog()==DialogResult.OK)
                {
                    _cfg=f.Result; Logger.Enabled=_cfg.EnableLogging; Config.Save(_cfg);
                    EnsureAutoStart(_cfg.AutoStart); StartOrRestart();
                }
            }
        }

        void EnsureAutoStart(bool enable)
        {
            try{
                using(var run=Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run",true))
                { if(run==null) return; const string name="MirrorAudio"; if(enable) run.SetValue(name,"\""+Application.ExecutablePath+"\""); else run.DeleteValue(name,false); }
            }catch{}
        }

        // —— 设备事件：去抖重启 —— //
        public void OnDeviceStateChanged(string id,DeviceState st){ if(IsRelevant(id)) Debounce(); }
        public void OnDeviceAdded(string id){ if(IsRelevant(id)) Debounce(); }
        public void OnDeviceRemoved(string id){ if(IsRelevant(id)) Debounce(); }
        public void OnDefaultDeviceChanged(DataFlow f,Role r,string id){ if(f==DataFlow.Render && r==Role.Multimedia && string.IsNullOrEmpty(_cfg!=null?_cfg.InputDeviceId:null)) Debounce(); }
        public void OnPropertyValueChanged(string id,PropertyKey key){ if(IsRelevant(id)) Debounce(); }
        bool IsRelevant(string id){ if(string.IsNullOrEmpty(id)||_cfg==null) return false; return id.Equals(_cfg.InputDeviceId,StringComparison.OrdinalIgnoreCase)||id.Equals(_cfg.MainDeviceId,StringComparison.OrdinalIgnoreCase)||id.Equals(_cfg.AuxDeviceId,StringComparison.OrdinalIgnoreCase); }
        void Debounce(){ _debounce.Stop(); _debounce.Start(); }

        // —— 主流程 —— //
        void StartOrRestart()
        {
            Stop();
            if(_mm==null) _mm=new MMDeviceEnumerator();

            _inDev=FindById(_cfg.InputDeviceId,DataFlow.Capture)??FindById(_cfg.InputDeviceId,DataFlow.Render);
            if(_inDev==null) _inDev=_mm.GetDefaultAudioEndpoint(DataFlow.Render,Role.Multimedia);
            _outMain=FindById(_cfg.MainDeviceId,DataFlow.Render);
            _outAux =FindById(_cfg.AuxDeviceId, DataFlow.Render);
            _inDevName=_inDev!=null?_inDev.FriendlyName:"-";
            if(_outMain==null||_outAux==null){ MessageBox.Show("请在“设置”选择主/副输出设备。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information); return; }

            WaveFormat inFmt;
            WaveFormat inputRequested=null, inputAccepted=null, inputMix=null;

            if(_inDev.DataFlow==DataFlow.Capture)
            {
                // 录音设备：沿用共享默认
                var cap=new WasapiCapture(_inDev){ ShareMode=AudioClientShareMode.Shared };
                _capture=cap; inFmt=cap.WaveFormat; _inRoleStr="录音";
                try{ inputMix=_inDev.AudioClient.MixFormat; }catch{}
                _inReqStr="录音-系统混音"; _inAccStr=Fmt(inFmt); _inMixStr=Fmt(inputMix);
            }
            else
            {
                // —— 环回设备：按策略请求指定格式（B 方案） —— //
                _inRoleStr="环回";

                var cap=new WasapiLoopbackCapture(_inDev);
                string negoLog="-";
                try{ inputMix=_inDev.AudioClient.MixFormat; }catch{}

                // 构造请求
                var req = new InputFormatRequest {
                    Strategy = _cfg.InputFormatStrategy,
                    CustomSampleRate = _cfg.InputCustomSampleRate,
                    CustomBitDepth = _cfg.InputCustomBitDepth,
                    Channels = 2
                };

                var wf = InputFormatHelper.NegotiateLoopbackFormat(_inDev, req, out negoLog, out inputMix, out inputAccepted, out inputRequested);
                if (wf != null) cap.WaveFormat = wf;
                _capture=cap; inFmt=cap.WaveFormat;

                _inReqStr = InputFormatHelper.Fmt(inputRequested);
                _inAccStr = InputFormatHelper.Fmt(inputAccepted ?? inFmt);
                _inMixStr = InputFormatHelper.Fmt(inputMix);
                if(Logger.Enabled) Logger.Info("Loopback negotiation:\r\n"+negoLog);
            }

            _inFmtStr=Fmt(inFmt);
            if(Logger.Enabled) Logger.Info("Input: "+_inDev.FriendlyName+" | "+_inFmtStr+" | "+_inRoleStr);

            _bufMain=new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true, BufferDuration=TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs*8,120)) };
            _bufAux =new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true, BufferDuration=TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs *6,150)) };

            GetPeriods(_outMain,out _defMainMs,out _minMainMs); GetPeriods(_outAux,out _defAuxMs,out _minAuxMs);

            // —— 主通道：独占/共享 + 事件/轮询 —— //
            _srcMain=_bufMain; _resMain=null; _mainIsExclusive=false; _mainEventSyncUsed=false; _mainBufEffectiveMs=_cfg.MainBufMs; _mainFmtStr="-";
            _mainNoSRC=false; _mainResampling=false;
            var desiredMain=new WaveFormat(_cfg.MainRate,_cfg.MainBits,2);
            bool isLoopMain=(_inDev.DataFlow==DataFlow.Render)&&_inDev.ID==_outMain.ID;
            bool wantExMain=(_cfg.MainShare==ShareModeOption.Exclusive||_cfg.MainShare==ShareModeOption.Auto)&&!isLoopMain;
            if(isLoopMain&&(_cfg.MainShare!=ShareModeOption.Shared)) MessageBox.Show("输入为主设备环回，独占冲突，主通道改走共享。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information);

            WaveFormat mainTargetFmt=null;

            if(wantExMain)
            {
                if(SupportsExclusive(_outMain,desiredMain)){
                    bool needRateChange = (inFmt.SampleRate!=desiredMain.SampleRate) || (inFmt.Channels!=desiredMain.Channels);
                    if(needRateChange) _srcMain=_resMain=new MediaFoundationResampler(_bufMain,desiredMain){ ResamplerQuality=50 };
                    int ms=Buf(_cfg.MainBufMs,true,_defMainMs,_minMainMs);
                    _mainOut=CreateOut(_outMain,AudioClientShareMode.Exclusive,_cfg.MainSync,ms,_srcMain,out _mainEventSyncUsed);
                    if(_mainOut!=null){ _mainIsExclusive=true; _mainBufEffectiveMs=ms; _mainFmtStr=Fmt(desiredMain); mainTargetFmt=desiredMain; _mainResampling = needRateChange; _mainNoSRC = !needRateChange; }
                }
                if(_mainOut==null && _cfg.MainBits==24){
                    var fmt32=new WaveFormat(_cfg.MainRate,32,2);
                    if(SupportsExclusive(_outMain,fmt32)){
                        bool needRateChange = (inFmt.SampleRate!=fmt32.SampleRate) || (inFmt.Channels!=fmt32.Channels);
                        _srcMain=_resMain=new MediaFoundationResampler(_bufMain,fmt32){ ResamplerQuality=50 };
                        int ms=Buf(_cfg.MainBufMs,true,_defMainMs,_minMainMs);
                        _mainOut=CreateOut(_outMain,AudioClientShareMode.Exclusive,_cfg.MainSync,ms,_srcMain,out _mainEventSyncUsed);
                        if(_mainOut!=null){ _mainIsExclusive=true; _mainBufEffectiveMs=ms; _mainFmtStr=Fmt(fmt32); mainTargetFmt=fmt32; _mainResampling = needRateChange; _mainNoSRC = !needRateChange; }
                    }
                }
                if(_mainOut==null && _cfg.MainShare==ShareModeOption.Exclusive){ MessageBox.Show("主通道独占失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Warning); Cleanup(); return; }
            }
            if(_mainOut==null)
            {
                int ms=Buf(_cfg.MainBufMs,false,_defMainMs);
                _mainOut=CreateOut(_outMain,AudioClientShareMode.Shared,_cfg.MainSync,ms,_bufMain,out _mainEventSyncUsed);
                if(_mainOut==null){ MessageBox.Show("主通道初始化失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Error); Cleanup(); return; }
                _mainBufEffectiveMs=ms; try{ mainTargetFmt=_outMain.AudioClient.MixFormat; _mainFmtStr=Fmt(mainTargetFmt);}catch{ _mainFmtStr="系统混音"; }
                _mainResampling = (inFmt.SampleRate != (mainTargetFmt!=null?mainTargetFmt.SampleRate:inFmt.SampleRate) || inFmt.Channels != (mainTargetFmt!=null?mainTargetFmt.Channels:inFmt.Channels));
                _mainNoSRC = !_mainResampling;
            }

            // —— 副通道 —— //
            _srcAux=_bufAux; _resAux=null; _auxIsExclusive=false; _auxEventSyncUsed=false; _auxBufEffectiveMs=_cfg.AuxBufMs; _auxFmtStr="-";
            _auxNoSRC=false; _auxResampling=false;
            var desiredAux=new WaveFormat(_cfg.AuxRate,_cfg.AuxBits,2);
            bool isLoopAux=(_inDev.DataFlow==DataFlow.Render)&&_inDev.ID==_outAux.ID;
            bool wantExAux=(_cfg.AuxShare==ShareModeOption.Exclusive||_cfg.AuxShare==ShareModeOption.Auto)&&!isLoopAux;
            if(isLoopAux&&(_cfg.AuxShare!=ShareModeOption.Shared)) MessageBox.Show("输入为副设备环回，独占冲突，副通道改走共享。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information);

            WaveFormat auxTargetFmt=null;

            if(wantExAux)
            {
                if(SupportsExclusive(_outAux,desiredAux)){
                    bool needRateChange = (inFmt.SampleRate!=desiredAux.SampleRate) || (inFmt.Channels!=desiredAux.Channels);
                    if(needRateChange) _srcAux=_resAux=new MediaFoundationResampler(_bufAux,desiredAux){ ResamplerQuality=40 };
                    int ms=Buf(_cfg.AuxBufMs,true,_defAuxMs,_minAuxMs);
                    _auxOut=CreateOut(_outAux,AudioClientShareMode.Exclusive,_cfg.AuxSync,ms,_srcAux,out _auxEventSyncUsed);
                    if(_auxOut!=null){ _auxIsExclusive=true; _auxBufEffectiveMs=ms; _auxFmtStr=Fmt(desiredAux); auxTargetFmt=desiredAux; _auxResampling=needRateChange; _auxNoSRC=!needRateChange; }
                }
                if(_auxOut==null && _cfg.AuxBits==24){
                    var fmt32=new WaveFormat(_cfg.AuxRate,32,2);
                    if(SupportsExclusive(_outAux,fmt32)){
                        bool needRateChange = (inFmt.SampleRate!=fmt32.SampleRate) || (inFmt.Channels!=fmt32.Channels);
                        _srcAux=_resAux=new MediaFoundationResampler(_bufAux,fmt32){ ResamplerQuality=40 };
                        int ms=Buf(_cfg.AuxBufMs,true,_defAuxMs,_minAuxMs);
                        _auxOut=CreateOut(_outAux,AudioClientShareMode.Exclusive,_cfg.AuxSync,ms,_srcAux,out _auxEventSyncUsed);
                        if(_auxOut!=null){ _auxIsExclusive=true; _auxBufEffectiveMs=ms; _auxFmtStr=Fmt(fmt32); auxTargetFmt=fmt32; _auxResampling=needRateChange; _auxNoSRC=!needRateChange; }
                    }
                }
                if(_auxOut==null && _cfg.AuxShare==ShareModeOption.Exclusive){ MessageBox.Show("副通道独占失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Warning); Cleanup(); return; }
            }
            if(_auxOut==null)
            {
                int ms=Buf(_cfg.AuxBufMs,false,_defAuxMs);
                _auxOut=CreateOut(_outAux,AudioClientShareMode.Shared,_cfg.AuxSync,ms,_bufAux,out _auxEventSyncUsed);
                if(_auxOut==null){ MessageBox.Show("副通道初始化失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Error); Cleanup(); return; }
                _auxBufEffectiveMs=ms; try{ auxTargetFmt=_outAux.AudioClient.MixFormat; _auxFmtStr=Fmt(auxTargetFmt);}catch{ _auxFmtStr="系统混音"; }
                _auxResampling = (inFmt.SampleRate != (auxTargetFmt!=null?auxTargetFmt.SampleRate:inFmt.SampleRate) || inFmt.Channels != (auxTargetFmt!=null?auxTargetFmt.Channels:inFmt.Channels));
                _auxNoSRC = !_auxResampling;
            }

            _capture.DataAvailable+=OnIn; _capture.RecordingStopped+=OnStopRec;
            try{
                _mainOut.Play(); _auxOut.Play(); _capture.StartRecording(); _running=true;
                if(Logger.Enabled){
                    Logger.Info("Main: "+(_mainIsExclusive?"独占":"共享")+" | "+(_mainEventSyncUsed?"事件":"轮询")+" | "+_mainBufEffectiveMs+"ms");
                    Logger.Info("Aux : "+(_auxIsExclusive ?"独占":"共享")+" | "+(_auxEventSyncUsed ?"事件":"轮询")+" | "+_auxBufEffectiveMs +"ms");
                }
            }
            catch(Exception ex){ Logger.Crash("Start",ex); MessageBox.Show("启动失败："+ex.Message,"MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Error); Stop(); }
        }

        void OnIn(object s, WaveInEventArgs e){ _bufMain.AddSamples(e.Buffer,0,e.BytesRecorded); _bufAux.AddSamples(e.Buffer,0,e.BytesRecorded); }
        void OnStopRec(object s, StoppedEventArgs e){ if(_bufMain!=null) _bufMain.ClearBuffer(); if(_bufAux!=null) _bufAux.ClearBuffer(); }

        public void Stop()
        {
            if(!_running){ DisposeAll(); return; }
            try{ _capture?.StopRecording(); }catch{} try{ _mainOut?.Stop(); }catch{} try{ _auxOut?.Stop(); }catch{}
            Thread.Sleep(20); DisposeAll(); _running=false; _tray.ShowBalloonTip(600,"MirrorAudio","已停止",ToolTipIcon.Info);
        }

        void DisposeAll()
        {
            try{ if(_capture!=null){ _capture.DataAvailable-=OnIn; _capture.RecordingStopped-=OnStopRec; _capture.Dispose(); } }catch{} _capture=null;
            try{ _mainOut?.Dispose(); }catch{} _mainOut=null;
            try{ _auxOut ?.Dispose(); }catch{} _auxOut =null;
            try{ _resMain?.Dispose(); }catch{} _resMain=null;
            try{ _resAux ?.Dispose(); }catch{} _resAux =null;
            _bufMain=null; _bufAux=null;
        }
        void Cleanup(){ DisposeAll(); }

        // —— 状态给设置窗体 —— //
        public StatusSnapshot GetStatusSnapshot()
        {
            
            // 计算对齐倍数：独占优先按最小周期倍数；否则按默认周期倍数
            double mainMul = 0, auxMul = 0;
            try {
                if (_mainBufEffectiveMs > 0) {
                    double baseMs = (_mainIsExclusive && _minMainMs > 0) ? _minMainMs : _defMainMs;
                    if (baseMs > 0) mainMul = Math.Round(_mainBufEffectiveMs / baseMs, 2);
                }
                if (_auxBufEffectiveMs > 0) {
                    double baseMs = (_auxIsExclusive && _minAuxMs > 0) ? _minAuxMs : _defAuxMs;
                    if (baseMs > 0) auxMul = Math.Round(_auxBufEffectiveMs / baseMs, 2);
                }
            } catch {}
return new StatusSnapshot{
                Running=_running,
                InputRole=_inRoleStr, InputFormat=_inFmtStr, InputDevice=_inDevName,
                InputRequested=_inReqStr, InputAccepted=_inAccStr, InputMix=_inMixStr,
                MainDevice=_outMain!=null?_outMain.FriendlyName:SafeName(_cfg.MainDeviceId,DataFlow.Render),
                AuxDevice =_outAux !=null?_outAux .FriendlyName:SafeName(_cfg.AuxDeviceId ,DataFlow.Render),
                MainMode=_mainOut!=null?(_mainIsExclusive?"独占":"共享"):"-", AuxMode=_auxOut!=null?(_auxIsExclusive?"独占":"共享"):"-",
                MainSync=_mainOut!=null?(_mainEventSyncUsed?"事件":"轮询"):"-",  AuxSync=_auxOut!=null?(_auxEventSyncUsed?"事件":"轮询"):"-",
                MainFormat=_mainOut!=null?_mainFmtStr:"-", AuxFormat=_auxOut!=null?_auxFmtStr:"-",
                MainBufferRequestedMs=_cfg.MainBufMs, AuxBufferRequestedMs=_cfg.AuxBufMs,
                MainBufferMs=_mainOut!=null?_mainBufEffectiveMs:0, AuxBufferMs=_auxOut!=null?_auxBufEffectiveMs:0,
                MainDefaultPeriodMs=_defMainMs, MainMinimumPeriodMs=_minMainMs, AuxDefaultPeriodMs=_defAuxMs, AuxMinimumPeriodMs=_minAuxMs,
                MainAlignedMultiple=mainMul, AuxAlignedMultiple=auxMul,
                MainNoSRC=_mainNoSRC, AuxNoSRC=_auxNoSRC, MainResampling=_mainResampling, AuxResampling=_auxResampling,
                MainBufferMultiple=(_mainBufEffectiveMs>0 && _minMainMs>0)? _mainBufEffectiveMs/_minMainMs:0,
                AuxBufferMultiple =(_auxBufEffectiveMs >0 && _minAuxMs >0)? _auxBufEffectiveMs /_minAuxMs:0
            };
        }

        string SafeName(string id, DataFlow flow)
        {
            if(string.IsNullOrEmpty(id)) return "-";
            try{ foreach(var d in _mm.EnumerateAudioEndPoints(flow,DeviceState.Active)) if(d.ID==id) return d.FriendlyName; }catch{}
            return "-";
        }
        MMDevice FindById(string id,DataFlow flow)
        {
            if(string.IsNullOrEmpty(id)) return null;
            try{ foreach(var d in _mm.EnumerateAudioEndPoints(flow,DeviceState.Active)) if(d.ID==id) return d; }catch{}
            return null;
        }
        
        static string Fmt(WaveFormat wf)
        {
            if (wf == null) return "-";
            // 容器位深
            string containerBits = (wf.Encoding == WaveFormatEncoding.IeeeFloat) ? "32" : wf.BitsPerSample.ToString();
            // 线缆/有效位深：优先识别 float，其次通过反射尝试 /*ValidBitsPerSample*/，最后退回容器位深
            string effectiveBits = (wf.Encoding == WaveFormatEncoding.IeeeFloat) ? "32f" : containerBits;
            try {
                var prop = wf.GetType().GetProperty("/*ValidBitsPerSample*/");
                if (prop != null) {
                    var vObj = prop.GetValue(wf, null);
                    if (vObj != null) {
                        int v = Convert.ToInt32(vObj);
                        if (v > 0) effectiveBits = v.ToString();
                    }
                }
            } catch {}
            return $"{wf.SampleRate}Hz/{containerBits}bit→{effectiveBits}bit/{wf.Channels}ch";
        }
