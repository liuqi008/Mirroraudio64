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
    static class Logger
    {
        public static bool Enabled; // 默认关闭
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        static readonly string CrashPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
        public static void Info(string s){ if(!Enabled) return; try{ File.AppendAllText(LogPath,"["+DateTime.Now.ToString("HH:mm:ss")+"] "+s+"\r\n"); }catch{} }
        public static void Crash(string where, Exception ex){ if(ex==null) return; try{ File.AppendAllText(CrashPath,$"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n"); }catch{} }
    }

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

    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    [DataContract]
    public sealed class AppSettings
    {
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare=ShareModeOption.Auto, AuxShare=ShareModeOption.Shared;
        [DataMember] public SyncModeOption  MainSync =SyncModeOption.Auto,  AuxSync =SyncModeOption.Auto;
        [DataMember] public int MainRate=192000, MainBits=24, MainBufMs=12;
        [DataMember] public int AuxRate =48000,  AuxBits =16,  AuxBufMs =150;
        [DataMember] public bool AutoStart=false, EnableLogging=false;
    }

    public sealed class StatusSnapshot
    {
        public bool Running; public string InputRole,InputFormat,InputDevice,MainDevice,AuxDevice,MainMode,AuxMode,MainSync,AuxSync,MainFormat,AuxFormat;
        public int MainBufferMs,AuxBufferMs; public double MainDefaultPeriodMs,MainMinimumPeriodMs,AuxDefaultPeriodMs,AuxMinimumPeriodMs;
    
        public bool MainPassthrough, AuxPassthrough;
}

    static class Config
    {
        static readonly string Dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"MirrorAudio");
        static readonly string FilePath=Path.Combine(Dir,"settings.json");
        public static AppSettings Load(){ try{ if(!File.Exists(FilePath)) return new AppSettings(); using(var fs=File.OpenRead(FilePath)) return (AppSettings)new DataContractJsonSerializer(typeof(AppSettings)).ReadObject(fs);}catch{ return new AppSettings(); }}
        public static void Save(AppSettings s){ try{ if(!Directory.Exists(Dir)) Directory.CreateDirectory(Dir); using(var fs=File.Create(FilePath)) new DataContractJsonSerializer(typeof(AppSettings)).WriteObject(fs,s);}catch{} }
    }

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
            if(_inDev.DataFlow==DataFlow.Capture){ var cap=new WasapiCapture(_inDev){ ShareMode=AudioClientShareMode.Shared }; _capture=cap; inFmt=cap.WaveFormat; _inRoleStr="录音"; }
            else{ var cap=new WasapiLoopbackCapture(_inDev); _capture=cap; inFmt=cap.WaveFormat; _inRoleStr="环回"; }
            _inFmtStr=Fmt(inFmt); if(Logger.Enabled) Logger.Info("Input: "+_inDev.FriendlyName+" | "+_inFmtStr+" | "+_inRoleStr);

            _bufMain=new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true, BufferDuration=TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs*8,120)) };
            _bufAux =new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow=true, ReadFully=true, BufferDuration=TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs *6,150)) };

            GetPeriods(_outMain,out _defMainMs,out _minMainMs); GetPeriods(_outAux,out _defAuxMs,out _minAuxMs);

            // 主通道：独占/共享 + 事件/轮询
            _srcMain=_bufMain; _resMain=null; _mainIsExclusive=false; _mainEventSyncUsed=false; _mainBufEffectiveMs=0; _mainFmtStr="-";
            var desiredMain=new WaveFormat(_cfg.MainRate,_cfg.MainBits,2);
            bool isLoopMain=(_inDev.DataFlow==DataFlow.Render)&&_inDev.ID==_outMain.ID;
            bool wantExMain=(_cfg.MainShare==ShareModeOption.Exclusive||_cfg.MainShare==ShareModeOption.Auto)&&!isLoopMain;
            if(isLoopMain&&(_cfg.MainShare!=ShareModeOption.Shared)) MessageBox.Show("输入为主设备环回，独占冲突，主通道改走共享。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information);

            if(wantExMain)
            {
                if(SupportsExclusive(_outMain,desiredMain)){ if(!Eq(inFmt,desiredMain)) _srcMain=_resMain=new MediaFoundationResampler(_bufMain,desiredMain){ ResamplerQuality=50 };
                    int ms=Buf(_cfg.MainBufMs,true,_defMainMs,_minMainMs);
                    _mainOut=CreateOut(_outMain,AudioClientShareMode.Exclusive,_cfg.MainSync,ms,_srcMain,out _mainEventSyncUsed);
                    if(_mainOut!=null){ _mainIsExclusive=true; _mainBufEffectiveMs=CalcEffectiveMs(_mainOut); _mainFmtStr=Fmt(desiredMain); }
                }
                if(_mainOut==null && _cfg.MainBits==24){ var fmt32=new WaveFormat(_cfg.MainRate,32,2); if(SupportsExclusive(_outMain,fmt32)){
                        _srcMain=_resMain=new MediaFoundationResampler(_bufMain,fmt32){ ResamplerQuality=50 };
                        int ms=Buf(_cfg.MainBufMs,true,_defMainMs,_minMainMs);
                        _mainOut=CreateOut(_outMain,AudioClientShareMode.Exclusive,_cfg.MainSync,ms,_srcMain,out _mainEventSyncUsed);
                        if(_mainOut!=null){ _mainIsExclusive=true; _mainBufEffectiveMs=CalcEffectiveMs(_mainOut); _mainFmtStr=Fmt(fmt32); }
                }}
                if(_mainOut==null && _cfg.MainShare==ShareModeOption.Exclusive){ MessageBox.Show("主通道独占失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Warning); Cleanup(); return; }
            }
            if(_mainOut==null)
            {
                int ms=Buf(_cfg.MainBufMs,false,_defMainMs);
                _mainOut=CreateOut(_outMain,AudioClientShareMode.Shared,_cfg.MainSync,ms,_bufMain,out _mainEventSyncUsed);
                if(_mainOut==null){ MessageBox.Show("主通道初始化失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Error); Cleanup(); return; }
                _mainBufEffectiveMs=CalcEffectiveMs(_mainOut); try{ _mainFmtStr=Fmt(_outMain.AudioClient.MixFormat);}catch{ _mainFmtStr="系统混音"; }
            }

            // 副通道
            _srcAux=_bufAux; _resAux=null; _auxIsExclusive=false; _auxEventSyncUsed=false; _auxBufEffectiveMs=0; _auxFmtStr="-";
            var desiredAux=new WaveFormat(_cfg.AuxRate,_cfg.AuxBits,2);
            bool isLoopAux=(_inDev.DataFlow==DataFlow.Render)&&_inDev.ID==_outAux.ID;
            bool wantExAux=(_cfg.AuxShare==ShareModeOption.Exclusive||_cfg.AuxShare==ShareModeOption.Auto)&&!isLoopAux;
            if(isLoopAux&&(_cfg.AuxShare!=ShareModeOption.Shared)) MessageBox.Show("输入为副设备环回，独占冲突，副通道改走共享。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information);

            if(wantExAux)
            {
                if(SupportsExclusive(_outAux,desiredAux)){ if(!Eq(inFmt,desiredAux)) _srcAux=_resAux=new MediaFoundationResampler(_bufAux,desiredAux){ ResamplerQuality=40 };
                    int ms=Buf(_cfg.AuxBufMs,true,_defAuxMs,_minAuxMs);
                    _auxOut=CreateOut(_outAux,AudioClientShareMode.Exclusive,_cfg.AuxSync,ms,_srcAux,out _auxEventSyncUsed);
                    if(_auxOut!=null){ _auxIsExclusive=true; _auxBufEffectiveMs=CalcEffectiveMs(_auxOut); _auxFmtStr=Fmt(desiredAux); }
                }
                if(_auxOut==null && _cfg.AuxBits==24){ var fmt32=new WaveFormat(_cfg.AuxRate,32,2); if(SupportsExclusive(_outAux,fmt32)){
                        _srcAux=_resAux=new MediaFoundationResampler(_bufAux,fmt32){ ResamplerQuality=40 };
                        int ms=Buf(_cfg.AuxBufMs,true,_defAuxMs,_minAuxMs);
                        _auxOut=CreateOut(_outAux,AudioClientShareMode.Exclusive,_cfg.AuxSync,ms,_srcAux,out _auxEventSyncUsed);
                        if(_auxOut!=null){ _auxIsExclusive=true; _auxBufEffectiveMs=CalcEffectiveMs(_auxOut); _auxFmtStr=Fmt(fmt32); }
                }}
                if(_auxOut==null && _cfg.AuxShare==ShareModeOption.Exclusive){ MessageBox.Show("副通道独占失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Warning); Cleanup(); return; }
            }
            if(_auxOut==null)
            {
                int ms=Buf(_cfg.AuxBufMs,false,_defAuxMs);
                _auxOut=CreateOut(_outAux,AudioClientShareMode.Shared,_cfg.AuxSync,ms,_bufAux,out _auxEventSyncUsed);
                if(_auxOut==null){ MessageBox.Show("副通道初始化失败。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Error); Cleanup(); return; }
                _auxBufEffectiveMs=CalcEffectiveMs(_auxOut); try{ _auxFmtStr=Fmt(_outAux.AudioClient.MixFormat);}catch{ _auxFmtStr="系统混音"; }
            }

            _capture.DataAvailable+=OnIn; _capture.RecordingStopped+=OnStopRec;
            try{ _mainOut.Play(); _auxOut.Play(); _capture.StartRecording(); _running=true;
                 if(Logger.Enabled){ Logger.Info("Main: "+(_mainIsExclusive?"独占":"共享")+" | "+(_mainEventSyncUsed?"事件":"轮询")+" | "+_mainBufEffectiveMs+"ms");
                                     Logger.Info("Aux : "+(_auxIsExclusive ?"独占":"共享")+" | "+(_auxEventSyncUsed ?"事件":"轮询")+" | "+_auxBufEffectiveMs +"ms"); } }
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
            return new StatusSnapshot{
                Running=_running, InputRole=_inRoleStr, InputFormat=_inFmtStr, InputDevice=_inDevName,
                MainDevice=_outMain!=null?_outMain.FriendlyName:SafeName(_cfg.MainDeviceId,DataFlow.Render),
                AuxDevice =_outAux !=null?_outAux .FriendlyName:SafeName(_cfg.AuxDeviceId ,DataFlow.Render),
                MainMode=_mainOut!=null?(_mainIsExclusive?"独占":"共享"):"-", AuxMode=_auxOut!=null?(_auxIsExclusive?"独占":"共享"):"-",
                MainSync=_mainOut!=null?(_mainEventSyncUsed?"事件":"轮询"):"-",  AuxSync=_auxOut!=null?(_auxEventSyncUsed?"事件":"轮询"):"-",
                MainFormat=_mainOut!=null?_mainFmtStr:"-", AuxFormat=_auxOut!=null?_auxFmtStr:"-",
                MainBufferMs=_mainOut!=null?_mainBufEffectiveMs:0, AuxBufferMs=_auxOut!=null?_auxBufEffectiveMs:0,
                MainDefaultPeriodMs=_defMainMs, MainMinimumPeriodMs=_minMainMs, AuxDefaultPeriodMs=_defAuxMs, AuxMinimumPeriodMs=_minAuxMs
            ,
                MainPassthrough = _mainOut!=null && _mainIsExclusive && _resMain==null,
                AuxPassthrough  = _auxOut !=null && _auxIsExclusive  && _resAux ==null
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
        static string Fmt(WaveFormat wf){ return wf==null?"-":(wf.SampleRate+"Hz/"+wf.BitsPerSample+"bit/"+wf.Channels+"ch"); }

        void GetPeriods(MMDevice dev,out double defMs,out double minMs)
        {
            defMs=10; minMs=2; if(dev==null) return; var id=dev.ID; Tuple<double,double> t;
            if(_periodCache.TryGetValue(id,out t)){ defMs=t.Item1; minMs=t.Item2; return; }
            try{
                long d100=0,m100=0; var ac=dev.AudioClient; var pD=ac.GetType().GetProperty("DefaultDevicePeriod"); var pM=ac.GetType().GetProperty("MinimumDevicePeriod");
                if(pD!=null){ var v=pD.GetValue(ac,null); if(v!=null) d100=Convert.ToInt64(v); }
                if(pM!=null){ var v=pM.GetValue(ac,null); if(v!=null) m100=Convert.ToInt64(v); }
                if(d100>0) defMs=d100/10000.0; if(m100>0) minMs=m100/10000.0;
            }catch{}
            _periodCache[id]=Tuple.Create(defMs,minMs);
        }

        static bool SupportsExclusive(MMDevice d,WaveFormat f){ try{ return d.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive,f);}catch{ return false; } }
        static bool Eq(WaveFormat a,WaveFormat b){ return a!=null&&b!=null&&a.SampleRate==b.SampleRate&&a.BitsPerSample==b.BitsPerSample&&a.Channels==b.Channels; }
        static int Buf(int want,bool exclusive,double defMs,double minMs=0)
        {
            int ms=want;
            if(exclusive){ int floor=(int)Math.Ceiling(defMs*3.0); if(ms<floor) ms=floor; if(minMs>0){ double k=Math.Ceiling(ms/minMs); ms=(int)Math.Ceiling(k*minMs); } }
            else{ int floor=(int)Math.Ceiling(defMs*2.0); if(ms<floor) ms=floor; }
            return ms;
        }

        WasapiOut CreateOut(MMDevice dev,AudioClientShareMode mode,SyncModeOption pref,int bufMs,IWaveProvider src,out bool eventUsed)
        {
            eventUsed=false; WasapiOut w;
            if(pref==SyncModeOption.Polling) return TryOut(dev,mode,false,bufMs,src);
            if(pref==SyncModeOption.Event){ w=TryOut(dev,mode,true,bufMs,src); if(w!=null){ eventUsed=true; return w; } return TryOut(dev,mode,false,bufMs,src); }
            w=TryOut(dev,mode,true,bufMs,src); if(w!=null){ eventUsed=true; return w; } return TryOut(dev,mode,false,bufMs,src);
        }
        WasapiOut TryOut(MMDevice dev,AudioClientShareMode mode,bool ev,int ms,IWaveProvider src)
        {
            try{ var w=new WasapiOut(dev,mode,ev,ms); w.Init(src); if(Logger.Enabled) Logger.Info($"OK {dev.FriendlyName} | {mode} | {(ev?"event":"poll")} | {ms}ms"); return w; }
            catch(Exception ex){ if(Logger.Enabled) Logger.Info($"Fail {dev.FriendlyName} | {mode} | {(ev?"event":"poll")} | {ms}ms | 0x{((uint)ex.HResult):X8} {ex.Message}"); return null; }
        }

        
        static int CalcEffectiveMs(WasapiOut w)
        {
            try
            {
                var ac = w?.AudioClient;
                if (ac == null) return 0;
                int frames = ac.AudioBufferSize;
                var wf = ac.MixFormat;
                if (wf == null || wf.SampleRate <= 0) return 0;
                double ms = (frames * 1000.0) / wf.SampleRate;
                return (int)Math.Round(ms, MidpointRounding.AwayFromZero);
            }
            catch { return 0; }
        }

        public void Dispose(){ Stop(); try{ _mm?.UnregisterEndpointNotificationCallback(this);}catch{} _debounce.Dispose(); _tray.Visible=false; _tray.Dispose(); _menu.Dispose(); }
    }
}
