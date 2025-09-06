using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces; // 热插拔事件
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace MirrorAudio
{
    static class Program
    {
        static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            bool createdNew;
            _mutex = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out createdNew);
            if (!createdNew) return;

            Application.ThreadException += (s, e) => SafeLog("UI", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => SafeLog("Non-UI", e.ExceptionObject as Exception);

            try { MediaFoundationApi.Startup(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) Application.Run();
            try { MediaFoundationApi.Shutdown(); } catch { }
        }

        static void SafeLog(string where, Exception ex)
        {
            if (ex == null) return;
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "MirrorAudio.crash.log");
                File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {where}: {ex}\r\n");
            } catch { }
        }
    }

    // —— 配置 —— //
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

        [DataMember] public ShareModeOption MainShare = ShareModeOption.Auto; // 主通道：共享/独占/自动
        [DataMember] public SyncModeOption  MainSync  = SyncModeOption.Auto;  // 主通道：事件/轮询/自动

        [DataMember] public int MainRate  = 192000;
        [DataMember] public int MainBits  = 24;
        [DataMember] public int MainBufMs = 12;   // 你要低延迟：12ms（轮询更稳）
        [DataMember] public int AuxBufMs  = 150;  // 副通道大缓冲、省资源

        [DataMember] public bool AutoStart = false;
        [DataMember] public bool EnableLogging = false;
    }

    static class Config
    {
        static string Dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MirrorAudio");
        static string FilePath = Path.Combine(Dir, "settings.json");

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

    // —— 主程序 —— //
    class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly string _log = Path.Combine(Path.GetTempPath(), "MirrorAudio.log");

        AppSettings _cfg = Config.Load();
        MMDeviceEnumerator _mm;
        System.Windows.Forms.Timer _debounce; // 400ms 去抖
        bool _running;

        // 设备 & 音频对象
        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture;
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux;
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain;

        // 记录最终模式
        bool _mainIsExclusive, _mainEventSyncUsed, _auxEventSyncUsed;
        int  _mainBufEffectiveMs, _auxBufEffectiveMs;

        public TrayApp()
        {
            _mm = new MMDeviceEnumerator();
            try { _mm.RegisterEndpointNotificationCallback(this); } catch { }

            _tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            _tray.Visible = true;
            _tray.Text = "MirrorAudio（三通道：输入/主/副）";

            var miStart = new ToolStripMenuItem("启动/重启(&S)", null, (s,e)=> StartOrRestart());
            var miStop  = new ToolStripMenuItem("停止(&T)", null, (s,e)=> Stop());
            var miSet   = new ToolStripMenuItem("设置(&G)...", null, (s,e)=> OpenSettings());
            var miLog   = new ToolStripMenuItem("打开日志目录", null, (s,e)=> Process.Start("explorer.exe", Path.GetTempPath()));
            var miExit  = new ToolStripMenuItem("退出(&X)", null, (s,e)=> { Stop(); Application.Exit(); });

            _menu.Items.AddRange(new ToolStripItem[]{ miStart, miStop, new ToolStripSeparator(), miSet, miLog, new ToolStripSeparator(), miExit });
            _tray.ContextMenuStrip = _menu;

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        // —— 菜单 —— //
        void OpenSettings()
        {
            using (var f = new SettingsForm(_cfg))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result;
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

        // —— 日志/提示 —— //
        void Log(string msg)
        {
            if (!_cfg.EnableLogging) return;
            try { File.AppendAllText(_log, "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + "\r\n"); } catch { }
        }
        static void Alert(string t, MessageBoxIcon ico)
        {
            MessageBox.Show(t, "MirrorAudio", MessageBoxButtons.OK, ico);
        }

        // —— 设备事件（事件驱动热插拔自愈） —— //
        public void OnDeviceStateChanged(string id, DeviceState st) { if (IsRelevant(id)) DebounceRestart($"state {st} @ {id}"); }
        public void OnDeviceAdded(string id)  { if (IsRelevant(id)) DebounceRestart($"added {id}"); }
        public void OnDeviceRemoved(string id){ if (IsRelevant(id)) DebounceRestart($"removed {id}"); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string id)
        {
            if (flow == DataFlow.Render && role == Role.Multimedia && string.IsNullOrEmpty(_cfg != null ? _cfg.InputDeviceId : null))
                DebounceRestart($"default render -> {id}");
        }
        public void OnPropertyValueChanged(string id, PropertyKey key) { if (IsRelevant(id)) DebounceRestart($"prop @ {id}"); }

        bool IsRelevant(string id)
        {
            if (string.IsNullOrEmpty(id) || _cfg == null) return false;
            return string.Equals(id, _cfg.InputDeviceId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.MainDeviceId,  StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.AuxDeviceId,   StringComparison.OrdinalIgnoreCase);
        }
        void DebounceRestart(string reason)
        {
            Log("Device change: " + reason + " → restart");
            if (_debounce == null)
            {
                _debounce = new System.Windows.Forms.Timer();
                _debounce.Interval = 400;
                _debounce.Tick += (s,e)=> { _debounce.Stop(); if (_menu.InvokeRequired) _menu.BeginInvoke((Action)StartOrRestart); else StartOrRestart(); };
            }
            _debounce.Stop(); _debounce.Start();
        }

        // —— 主流程 —— //
        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            // 输入：优先配置；否则默认渲染环回
            _inDev   = FindById(_mm, _cfg.InputDeviceId, DataFlow.Capture) ?? FindById(_mm, _cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_mm, _cfg.MainDeviceId, DataFlow.Render);
            _outAux  = FindById(_mm, _cfg.AuxDeviceId,  DataFlow.Render);

            if (_outMain == null || _outAux == null)
            {
                Alert("请先在“设置”里选择主/副输出设备。", MessageBoxIcon.Information);
                return;
            }

            // 输入管线：录音→WasapiCapture；渲染→WasapiLoopbackCapture
            WaveFormat inFmt;
            if (_inDev.DataFlow == DataFlow.Capture) { var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared }; _capture = cap; inFmt = cap.WaveFormat; }
            else                                      { var cap = new WasapiLoopbackCapture(_inDev); _capture = cap; inFmt = cap.WaveFormat; }
            Log("Input: " + _inDev.FriendlyName + $" | {inFmt.SampleRate}Hz/{inFmt.BitsPerSample}bit/{inFmt.Channels}ch");

            // 两路桥接缓冲
            _bufMain = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.MainBufMs * 8, 120)) };
            _bufAux  = new BufferedWaveProvider(inFmt){ DiscardOnBufferOverflow = true, ReadFully = true, BufferDuration = TimeSpan.FromMilliseconds(Math.Max(_cfg.AuxBufMs  * 6, 150)) };

            // 设备周期
            double defMain, minMain, defAux, minAux;
            GetDevicePeriodsMs(_outMain, out defMain, out minMain);
            GetDevicePeriodsMs(_outAux,  out defAux,  out minAux);
            Log($"Main DevicePeriod: default={defMain:0.##}ms, min={minMain:0.##}ms");
            Log($"Aux  DevicePeriod: default={defAux:0.##}ms,  min={minAux:0.##}ms");

            // —— 主通道：共享/独占 + 事件/轮询（可选） —— //
            _srcMain = _bufMain; _resMain = null;
            _mainIsExclusive = false; _mainEventSyncUsed = false; _mainBufEffectiveMs = _cfg.MainBufMs;

            WaveFormat desired = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
            bool forceExclusive = _cfg.MainShare == ShareModeOption.Exclusive;
            bool forceShared    = _cfg.MainShare == ShareModeOption.Shared;
            bool tryExclusive   = _cfg.MainShare == ShareModeOption.Auto;

            // 先试独占（若要求/自动）
            if (forceExclusive || tryExclusive)
            {
                if (IsFormatSupportedExclusive(_outMain, desired))
                {
                    if (!FormatsEqual(inFmt, desired))
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, desired){ ResamplerQuality = 50 };
                        _srcMain = _resMain;
                        Log($"Resampler: ON -> {desired.SampleRate}Hz/{desired.BitsPerSample}bit");
                    }
                    else Log("Resampler: OFF (input==device)");

                    int ms = SafeBuf(_cfg.MainBufMs, true, defMain);
                    _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Exclusive, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                    if (_mainOut != null) { _mainIsExclusive = true; _mainBufEffectiveMs = ms; }
                }
                else if (forceExclusive)
                {
                    Alert("“强制独占”失败：设备不支持该独占格式（请调整采样率/位深）。", MessageBoxIcon.Warning);
                    CleanupCreated(); return;
                }
            }

            // 若独占未成或强制共享，则走共享
            if (_mainOut == null)
            {
                _srcMain = _bufMain; // 共享交给系统混音/重采样
                Log("Resampler (shared): OFF");
                int ms = SafeBuf(_cfg.MainBufMs, false, defMain);
                _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, ms, _srcMain, out _mainEventSyncUsed);
                if (_mainOut == null) { Alert("主通道初始化失败（独占/共享均不可用）。", MessageBoxIcon.Error); CleanupCreated(); return; }
                _mainIsExclusive = false; _mainBufEffectiveMs = ms;
            }

            // —— 副通道（共享，事件优先，失败回退轮询） —— //
            _srcAux = _bufAux;
            int auxMs = Math.Max(_cfg.AuxBufMs, (int)Math.Ceiling(defAux * 2.0));
            _auxOut = CreateOutWithPolicy(_outAux, AudioClientShareMode.Shared, SyncModeOption.Auto, auxMs, _srcAux, out _auxEventSyncUsed);
            if (_auxOut == null) { Alert("副通道初始化失败（共享不可用）。", MessageBoxIcon.Error); CleanupCreated(); return; }
            _auxBufEffectiveMs = auxMs;

            // —— 绑定与启动 —— //
            _capture.DataAvailable += (s, e) => { _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded); _bufAux.AddSamples(e.Buffer, 0, e.BytesRecorded); };
            _capture.RecordingStopped += (s, e) => { if (_bufMain != null) _bufMain.ClearBuffer(); if (_bufAux != null) _bufAux.ClearBuffer(); };

            try
            {
                _mainOut.Play();
                _auxOut.Play();
                _capture.StartRecording();
                _running = true;

                string mainMode = _mainIsExclusive ? "独占" : "共享";
                string mainSync = _mainEventSyncUsed ? "事件" : "轮询";
                string auxSync  = _auxEventSyncUsed  ? "事件" : "轮询";
                Log($"Final Main: mode={mainMode}, sync={mainSync}, buffer={_mainBufEffectiveMs}ms");
                Log($"Final Aux : sync={auxSync}, buffer={_auxBufEffectiveMs}ms");
            }
            catch (Exception ex)
            {
                Log("Start failed: " + ex);
                Alert("启动失败：" + ex.Message, MessageBoxIcon.Error);
                Stop();
            }
        }

        public void Stop()
        {
            if (!_running)
            {
                DisposeAll(); return;
            }
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut.Stop();  } catch { }
            Thread.Sleep(40);
            DisposeAll();
            _running = false;
            _tray.ShowBalloonTip(900, "MirrorAudio", "已停止", ToolTipIcon.Info);
        }

        void DisposeAll()
        {
            try { if (_capture != null) _capture.Dispose(); } catch { } _capture = null;
            try { if (_mainOut != null) _mainOut.Dispose(); } catch { } _mainOut = null;
            try { if (_auxOut  != null) _auxOut .Dispose(); } catch { } _auxOut  = null;
            try { if (_resMain != null) _resMain.Dispose(); } catch { } _resMain = null;
            _bufMain = null; _bufAux = null;
        }

        // —— 工具函数 —— //
        static MMDevice FindById(MMDeviceEnumerator mm, string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return mm.EnumerateAudioEndPoints(flow, DeviceState.Active).FirstOrDefault(d => d.ID == id);
        }

        // 读取设备周期：优先属性（NAudio 2.x），备选反射方法，失败则兜底
        static void GetDevicePeriodsMs(MMDevice dev, out double defMs, out double minMs)
        {
            defMs = 10.0; minMs = 2.0;
            try
            {
                var ac = dev.AudioClient;
                long def100ns = 0, min100ns = 0;
                var pDef = ac.GetType().GetProperty("DefaultDevicePeriod");
                var pMin = ac.GetType().GetProperty("MinimumDevicePeriod");
                if (pDef != null) { object v = pDef.GetValue(ac, null); if (v != null) def100ns = Convert.ToInt64(v); }
                if (pMin != null) { object v = pMin.GetValue(ac, null); if (v != null) min100ns = Convert.ToInt64(v); }
                if (def100ns == 0 || min100ns == 0)
                {
                    var m = ac.GetType().GetMethod("GetDevicePeriod");
                    if (m != null)
                    {
                        object[] args = new object[] { 0L, 0L };
                        m.Invoke(ac, args);
                        def100ns = (long)args[0];
                        min100ns = (long)args[1];
                    }
                }
                if (def100ns > 0) defMs = def100ns / 10000.0;
                if (min100ns > 0) minMs = min100ns / 10000.0;
            } catch { }
        }

        static bool IsFormatSupportedExclusive(MMDevice dev, WaveFormat fmt)
        {
            try { return dev.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, fmt); }
            catch { return false; }
        }

        static bool FormatsEqual(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels;
        }

        static int SafeBuf(int desiredMs, bool exclusive, double defaultPeriodMs)
        {
            int minMs = (int)Math.Ceiling(defaultPeriodMs * (exclusive ? 3.0 : 2.0));
            return desiredMs < minMs ? minMs : desiredMs;
        }

        // —— 修复后的版本：不再在本地函数里写 out 参数 —— //
        WasapiOut CreateOutWithPolicy(MMDevice dev, AudioClientShareMode mode, SyncModeOption syncPref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false;
            WasapiOut w = null;

            // 强制轮询
            if (syncPref == SyncModeOption.Polling)
            {
                w = TryOut(dev, mode, false, bufMs, src);
                return w; // eventUsed 保持 false
            }

            // 强制事件：失败则回退轮询
            if (syncPref == SyncModeOption.Event)
            {
                w = TryOut(dev, mode, true, bufMs, src);
                if (w != null) { eventUsed = true; return w; }

                w = TryOut(dev, mode, false, bufMs, src);
                eventUsed = false;
                return w;
            }

            // Auto：事件优先，失败回退轮询
            w = TryOut(dev, mode, true, bufMs, src);
            if (w != null) { eventUsed = true; return w; }

            w = TryOut(dev, mode, false, bufMs, src);
            eventUsed = false;
            return w;
        }

        WasapiOut TryOut(MMDevice dev, AudioClientShareMode mode, bool eventSync, int bufMs, IWaveProvider src)
        {
            try
            {
                var w = new WasapiOut(dev, mode, eventSync, bufMs);
                w.Init(src);
                Log($"WasapiOut OK: {dev.FriendlyName} | mode={mode} event={eventSync} buf={bufMs}ms");
                return w;
            }
            catch (Exception ex)
            {
                Log($"WasapiOut failed: {dev.FriendlyName} | mode={mode} event={eventSync} buf={bufMs}ms | {ex.Message}");
                return null;
            }
        }

        void CleanupCreated()
        {
            try { if (_resMain != null) _resMain.Dispose(); } catch { }
            try { if (_mainOut != null) _mainOut.Dispose(); } catch { }
            try { if (_auxOut  != null) _auxOut  .Dispose(); } catch { }
            _resMain = null; _mainOut = null; _auxOut = null;
        }

        public void Dispose()
        {
            Stop();
            try { if (_mm != null) _mm.UnregisterEndpointNotificationCallback(this); } catch { }
            if (_debounce != null) _debounce.Dispose();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
        }
    }
}
