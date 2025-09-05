using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace MirrorAudio
{
    // 轻量日志：先判定再打印，避免无谓字符串分配
    static class Logger
    {
        public static bool Enabled;
        static readonly string _logPath = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");
        public static void Info(string msg)
        {
            if (!Enabled) return;
            try { File.AppendAllText(_logPath, "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + "\r\n"); } catch { }
        }
        public static void Crash(string where, Exception ex)
        {
            if (ex == null) return;
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
                File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n");
            } catch { }
        }
    }

    static class Program
    {
        static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            bool createdNew;
            _mutex = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out createdNew);
            if (!createdNew) return;

            Application.ThreadException += (s, e) => Logger.Crash("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Logger.Crash("Non-UI", e.ExceptionObject as Exception);

            try { MediaFoundationApi.Startup(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) Application.Run();
            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    // 配置类
    [DataContract]
    public enum ShareModeOption { [EnumMember] Auto, [EnumMember] Exclusive, [EnumMember] Shared }
    [DataContract]
    public enum SyncModeOption  { [EnumMember] Auto, [EnumMember] Event, [EnumMember] Polling }

    [DataContract]
    public class AppSettings
    {
        [DataMember] public string InputDeviceId;
        [DataMember] public string MainDeviceId;
        [DataMember] public string AuxDeviceId;
        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto;
        [DataMember] public SyncModeOption  MainSync  = SyncModeOption.Auto;
        [DataMember] public int MainRate  = 192000;
        [DataMember] public int MainBits  = 24;
        [DataMember] public int MainBufMs = 12;
        [DataMember] public ShareModeOption AuxShare = ShareModeOption.Shared;
        [DataMember] public SyncModeOption  AuxSync  = SyncModeOption.Auto;
        [DataMember] public int AuxRate  = 48000;   
        [DataMember] public int AuxBits  = 16;      
        [DataMember] public int AuxBufMs = 150;
        [DataMember] public bool AutoStart = false;
        [DataMember] public bool EnableLogging = false;
    }

    public class StatusSnapshot
    {
        public bool Running;
        public string InputRole;
        public string InputFormat;
        public string InputDevice;
        public string MainDevice;
        public string AuxDevice;
        public string MainMode;
        public string AuxMode;
        public string MainSync;
        public string AuxSync;
        public string MainFormat;
        public string AuxFormat;
        public int MainBufferMs;
        public int AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs;
        public double AuxDefaultPeriodMs,  AuxMinimumPeriodMs;
    }

    static class Config
    {
        static readonly string Dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
        static readonly string FilePath = Path.Combine(Dir, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new AppSettings();
                using (var fs = File.OpenRead(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    return (AppSettings)ser.ReadObject(fs);
                }
            } catch { return new AppSettings(); }
        }

        public static void Save(AppSettings s)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                using (var fs = File.Create(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    ser.WriteObject(fs, s);
                }
            } catch { }
        }
    }

    // 主程序
    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm;             
        readonly System.Windows.Forms.Timer _debounce;             
        bool _running;
        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture;
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux;
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain, _resAux;
        bool _mainIsExclusive, _mainEventSyncUsed, _auxIsExclusive, _auxEventSyncUsed;
        int  _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _inRoleStr = "-", _inFmtStr = "-", _inDevName = "-";
        string _mainFmtStr = "-", _auxFmtStr = "-";
        double _defMainMs, _minMainMs, _defAuxMs, _minAuxMs;
        readonly Dictionary<string, Tuple<double,double>> _periodCache = new Dictionary<string, Tuple<double,double>>(4);

        public TrayApp()
        {
            Logger.Enabled = _cfg.EnableLogging;
            _mm = new MMDeviceEnumerator();
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            try
            {
                if (File.Exists("MirrorAudio.ico")) _tray.Icon = new Icon("MirrorAudio.ico");
                else _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { _tray.Icon = SystemIcons.Application; }

            _tray.Visible = true;
            _tray.Text = "MirrorAudio";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, OnStartClick);
            var miStop  = new ToolStripMenuItem("停止(&T)", null, OnStopClick);
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, OnSettingsClick);
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s,e)=> Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)", null, (s,e)=> { Stop(); Application.Exit(); });

            _menu.Items.AddRange(new ToolStripItem[]{ miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            _debounce = new System.Windows.Forms.Timer();
            _debounce.Interval = 400;
            _debounce.Tick += OnDebounceTick;

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        public void Dispose()
        {
            // 释放资源
            DisposeAll();
        }

        // 实现 IMMNotificationClient 接口的方法
        public void OnDeviceStateChanged(string deviceId, DeviceState newState) { /* 处理设备状态变化 */ }
        public void OnDeviceAdded(string deviceId) { /* 处理设备添加 */ }
        public void OnDeviceRemoved(string deviceId) { /* 处理设备移除 */ }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string deviceId) { /* 处理默认设备变化 */ }
        public void OnPropertyValueChanged(string deviceId, PropertyKey propertyKey) { /* 处理属性变化 */ }

        void OnStartClick(object sender, EventArgs e) => StartOrRestart();
        void OnStopClick(object sender, EventArgs e)  => Stop();
        void OnSettingsClick(object sender, EventArgs e)
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
            } catch { }
        }

        public void Stop()
        {
            if (!_running)
            {
                DisposeAll();
                return;
            }
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut.Stop();  } catch { }
            Thread.Sleep(30);
            DisposeAll();
            _running = false;
            _tray.ShowBalloonTip(800, "MirrorAudio", "已停止", ToolTipIcon.Info);
        }

        void DisposeAll()
        {
            try { if (_capture != null) { _capture.DataAvailable -= OnCaptureDataAvailable; _capture.RecordingStopped -= OnRecordingStopped; _capture.Dispose(); } } catch { } _capture = null;
            try { if (_mainOut != null) _mainOut.Dispose(); } catch { } _mainOut = null;
            try { if (_auxOut  != null) _auxOut .Dispose(); } catch { } _auxOut  = null;
            try { if (_resMain != null) _resMain.Dispose(); } catch { } _resMain = null;
            try { if (_resAux  != null) _resAux .Dispose(); } catch { } _resAux  = null;
            _bufMain = null; _bufAux = null;
        }

        public StatusSnapshot GetStatusSnapshot()
        {
            var s = new StatusSnapshot
            {
                Running = _running,
                InputRole = _inRoleStr,
                InputFormat = _inFmtStr,
                InputDevice = _inDevName,
                MainDevice = _outMain != null ? _outMain.FriendlyName : SafeNameForId(_cfg.MainDeviceId, DataFlow.Render),
                AuxDevice  = _outAux  != null ? _outAux .FriendlyName : SafeNameForId(_cfg.AuxDeviceId,  DataFlow.Render),
                MainMode = _mainOut != null ? (_mainIsExclusive ? "独占" : "共享") : "-",
                AuxMode  = _auxOut  != null ? (_auxIsExclusive  ? "独占" : "共享") : "-",
                MainSync = _mainOut != null ? (_mainEventSyncUsed ? "事件" : "轮询") : "-",
                AuxSync  = _auxOut  != null ? (_auxEventSyncUsed  ? "事件" : "轮询") : "-",
                MainFormat = _mainOut != null ? _mainFmtStr : "-",
                AuxFormat  = _auxOut  != null ? _auxFmtStr  : "-",
                MainBufferMs = _mainOut != null ? _mainBufEffectiveMs : 0,
                AuxBufferMs  = _auxOut  != null ? _auxBufEffectiveMs  : 0,
                MainDefaultPeriodMs = _defMainMs,
                MainMinimumPeriodMs = _minMainMs,
                AuxDefaultPeriodMs  = _defAuxMs,
                AuxMinimumPeriodMs  = _minAuxMs
            };
            return s;
        }
    }
}
