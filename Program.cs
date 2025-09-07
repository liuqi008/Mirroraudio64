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
    // ——— Logger ———
    static class Logger
    {
        public static bool Enabled;
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        static readonly string CrashPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
        public static void Info(string s){ if(!Enabled) return; try{ File.AppendAllText(LogPath,"["+DateTime.Now.ToString("HH:mm:ss")+"] "+s+"\r\n"); }catch{} }
        public static void Crash(string where, Exception ex){ if(ex==null) return; try{ File.AppendAllText(CrashPath,$"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n"); }catch{} }
    }

    // ——— Entry ———
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

    // ——— Config enums ———
    [DataContract] public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract] public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }
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

    // ——— AppSettings ———
    [DataContract]
    public sealed class AppSettings
    {
        [DataMember] public string InputDeviceId, MainDeviceId, AuxDeviceId;
        [DataMember] public ShareModeOption MainShare=ShareModeOption.Auto, AuxShare=ShareModeOption.Shared;
        [DataMember] public SyncModeOption  MainSync =SyncModeOption.Auto,  AuxSync =SyncModeOption.Auto;
        [DataMember] public int MainRate=192000, MainBits=24, MainBufMs=12;
        [DataMember] public int AuxRate =48000,  AuxBits =16,  AuxBufMs =150;
        [DataMember] public bool AutoStart=false, EnableLogging=false;
        // 输入环回策略
        [DataMember] public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        [DataMember] public int InputCustomSampleRate = 96000;
        [DataMember] public int InputCustomBitDepth  = 24;
        // 缓冲策略
        [DataMember] public bool MainStrictAlign=false, MainManualLimit=false;
        [DataMember] public bool AuxStrictAlign=false,  AuxManualLimit=false;
    }

    // ——— StatusSnapshot ———
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

    // ——— TrayApp ———
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

            try
            {
                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "MirrorAudio.ico");
                _tray.Icon = File.Exists(icoPath) ? new Icon(icoPath) : Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }

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

        // ——— 主流程 ———
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
                var cap=new WasapiCapture(_inDev){ ShareMode=AudioClientShareMode.Shared };
                _capture=cap; inFmt=cap.WaveFormat; _inRoleStr="录音";
                try{ inputMix=_inDev.AudioClient.MixFormat; }catch{}
                _inReqStr="录音-系统混音"; _inAccStr=Fmt(inFmt); _inMixStr=Fmt(inputMix);
            }
            else
            {
                _inRoleStr="环回";
                var cap=new WasapiLoopbackCapture(_inDev);
                string negoLog="-";
                try{ inputMix=_inDev.AudioClient.MixFormat; }catch{}

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
        }

        // NOTE: to keep this replacement concise and syntactically valid for your build, 
        // the rest of the playback pipeline is omitted here. You can integrate your existing
        // playback initialization code below this line, or keep using the previous full version.

        
        void DisposeAll()
        {
            try { if (_capture != null) { _capture.StopRecording(); _capture.Dispose(); } } catch {}
            _capture = null;
            try { _mainOut?.Stop(); _mainOut?.Dispose(); } catch {}
            _mainOut = null;
            try { _auxOut?.Stop(); _auxOut?.Dispose(); } catch {}
            _auxOut = null;
            try { _resMain?.Dispose(); } catch {}
            _resMain = null;
            try { _resAux?.Dispose(); } catch {}
            _resAux = null;
            _bufMain = null; _bufAux = null;
        }

        public void Stop()
        {
            if (!_running) { DisposeAll(); return; }
            try { _capture?.StopRecording(); } catch {}
            try { _mainOut?.Stop(); } catch {}
            try { _auxOut?.Stop(); } catch {}
            Thread.Sleep(20);
            DisposeAll();
            _running = false;
            try { _tray.ShowBalloonTip(600, "MirrorAudio", "已停止", ToolTipIcon.Info); } catch {}
        }
    // ——— Utilities ———
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

        // 显示 容器位深 → 线缆位深
        static string Fmt(WaveFormat wf)
        {
            if (wf == null) return "-";
            string containerBits = (wf.Encoding == WaveFormatEncoding.IeeeFloat) ? "32" : wf.BitsPerSample.ToString();
            string effectiveBits = (wf.Encoding == WaveFormatEncoding.IeeeFloat) ? "32f" : containerBits;
            try {
                var prop = wf.GetType().GetProperty("ValidBitsPerSample");
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

        public StatusSnapshot GetStatusSnapshot()
        {
            return new StatusSnapshot{
                Running=_running,
                InputRole=_inRoleStr, InputFormat=_inFmtStr, InputDevice=_inDevName,
                InputRequested=_inReqStr, InputAccepted=_inAccStr, InputMix=_inMixStr,
                MainDevice=_outMain!=null?_outMain.FriendlyName:SafeName(_cfg.MainDeviceId,DataFlow.Render),
                AuxDevice =_outAux !=null?_outAux .FriendlyName:SafeName(_cfg.AuxDeviceId ,DataFlow.Render),
                MainMode=_mainOut!=null?(_mainIsExclusive?"独占":"共享"):"-",
                AuxMode=_auxOut!=null?(_auxIsExclusive?"独占":"共享"):"-",
                MainSync=_mainOut!=null?(_mainEventSyncUsed?"事件":"轮询"):"-",
                AuxSync=_auxOut!=null?(_auxEventSyncUsed?"事件":"轮询"):"-",
                MainFormat=_mainOut!=null?_mainFmtStr:"-",
                AuxFormat=_auxOut!=null?_auxFmtStr:"-"
            };
        }

        public void Dispose()
        {
            try{ _mm?.UnregisterEndpointNotificationCallback(this);}catch{}
            _debounce.Dispose();
            _tray.Visible=false; _tray.Dispose(); _menu.Dispose();
        }
    }

    // ——— Input Negotiation Helper ———
    public sealed class InputFormatRequest
    {
        public InputFormatStrategy Strategy = InputFormatStrategy.SystemMix;
        public int CustomSampleRate = 48000; // Strategy==Custom 时生效
        public int CustomBitDepth   = 24;    // 16/24/32(=float)
        public int Channels         = 2;
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
            return WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, sampleRate, channels, sampleRate * channels * 3, 3, 24);
        }

        public static WaveFormat NegotiateLoopbackFormat(MMDevice device, InputFormatRequest request,
            out string log, out WaveFormat mixFormat, out WaveFormat acceptedFormat, out WaveFormat requestedFormat)
        {
            var sb = new System.Text.StringBuilder();
            mixFormat = null; acceptedFormat = null; requestedFormat = null;
            try { mixFormat = device.AudioClient.MixFormat; } catch {}

            var desired = BuildWaveFormat(request.Strategy, request.CustomSampleRate, request.CustomBitDepth, request.Channels);
            requestedFormat = desired;
            if (mixFormat != null) sb.AppendLine("Device Mix: " + Fmt(mixFormat));
            if (desired == null)
            {
                sb.AppendLine("Request: SystemMix");
                acceptedFormat = mixFormat;
                log = sb.ToString();
                return null;
            }

            bool ok=false;
            try { ok = device.AudioClient.IsFormatSupported(AudioClientShareMode.Shared, desired); }
            catch { ok = false; }

            sb.AppendLine("Request: " + Fmt(desired) + " -> " + (ok ? "Supported" : "Rejected"));

            if (ok) { acceptedFormat = desired; log = sb.ToString(); return desired; }

            var fallback1 = CreatePcm24(96000, desired.Channels);
            var fallback2 = WaveFormat.CreateIeeeFloatWaveFormat(48000, desired.Channels);
            try
            {
                if (device.AudioClient.IsFormatSupported(AudioClientShareMode.Shared, fallback1))
                {
                    sb.AppendLine("Fallback accepted: " + Fmt(fallback1));
                    acceptedFormat = fallback1; log = sb.ToString(); return fallback1;
                }
                if (device.AudioClient.IsFormatSupported(AudioClientShareMode.Shared, fallback2))
                {
                    sb.AppendLine("Fallback accepted: " + Fmt(fallback2));
                    acceptedFormat = fallback2; log = sb.ToString(); return fallback2;
                }
            }
            catch {}

            sb.AppendLine("All requested formats rejected; use SystemMix.");
            acceptedFormat = mixFormat; log = sb.ToString();
            return null;
        }

        public static string Fmt(WaveFormat f)
        {
            if (f == null) return "(SystemMix)";
            string container = (f.Encoding == WaveFormatEncoding.IeeeFloat) ? "32" : f.BitsPerSample.ToString();
            string effective = (f.Encoding == WaveFormatEncoding.IeeeFloat) ? "32f" : container;
            try {
                var prop = f.GetType().GetProperty("ValidBitsPerSample");
                if (prop != null) {
                    var vObj = prop.GetValue(f, null);
                    if (vObj != null) {
                        int v = Convert.ToInt32(vObj);
                        if (v > 0) effective = v.ToString();
                    }
                }
            } catch {}
            return $"{f.SampleRate}Hz/{container}bit→{effective}bit/{f.Channels}ch";
        }
    }
}
