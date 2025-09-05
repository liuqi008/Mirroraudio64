using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace MirrorAudio
{
    sealed class TrayApp : IDisposable, IMMNotificationClient
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip { ShowCheckMargin = false, ShowImageMargin = true };
        readonly WinFormsTimer _debounce = new WinFormsTimer { Interval = 420 };

        AppSettings _cfg;
        MMDevice _inDev, _outMain, _outAux;
        IWaveIn _capture;
        WasapiOut _mainOut, _auxOut;
        BufferedWaveProvider _bufMain, _bufAux;

        bool _running;
        bool _mainIsExclusive, _auxIsExclusive;

        public TrayApp()
        {
            _cfg = Config.Load();
            Logger.Enabled = _cfg.EnableLogging;
            _tray.Icon = File.Exists("MirrorAudio.ico") ? new Icon("MirrorAudio.ico") : SystemIcons.Application;
            _tray.Visible = true;
            _tray.Text = "MirrorAudio";
            _menu.Items.AddRange(new[]
            {
                new ToolStripMenuItem("启动/重启(&S)", To16(SystemIcons.Information), OnStartClick),
                new ToolStripMenuItem("停止(&T)", To16(SystemIcons.Hand), OnStopClick),
                new ToolStripMenuItem("设置(&G)...", To16(SystemIcons.Information), OnSettingsClick),
                new ToolStripMenuItem("日志目录", To16(SystemIcons.Information), OnLogDirClick),
                new ToolStripMenuItem("退出(&X)", To16(SystemIcons.Error), OnExitClick)
            });
            _tray.ContextMenuStrip = _menu;

            _debounce.Tick += OnDebounceTick;

            EnsureAutoStart(_cfg.AutoStart);
            StartOrRestart();
        }

        static Bitmap To16(Icon ico)
        {
            try
            {
                using (var bmp = ico.ToBitmap())
                {
                    var img = new Bitmap(16, 16);
                    using (var g = Graphics.FromImage(img)) g.DrawImage(bmp, new Rectangle(0, 0, 16, 16));
                    return img;
                }
            } catch { return null; }
        }

        void OnStartClick(object sender, EventArgs e) => StartOrRestart();
        void OnStopClick(object sender, EventArgs e) => Stop();
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

        void OnLogDirClick(object sender, EventArgs e)
        {
            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (Directory.Exists(logDir)) System.Diagnostics.Process.Start(logDir);
            else MessageBox.Show("日志目录不存在。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void OnExitClick(object sender, EventArgs e) { Stop(); Application.Exit(); }

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

        public void OnDeviceStateChanged(string id, DeviceState st) { if (IsRelevant(id)) DebounceRestart(); }
        public void OnDeviceAdded(string id)  { if (IsRelevant(id)) DebounceRestart(); }
        public void OnDeviceRemoved(string id) { if (IsRelevant(id)) DebounceRestart(); }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string id)
        {
            if (flow == DataFlow.Render && role == Role.Multimedia && string.IsNullOrEmpty(_cfg != null ? _cfg.InputDeviceId : null))
                DebounceRestart();
        }
        public void OnPropertyValueChanged(string id, PropertyKey key) { if (IsRelevant(id)) DebounceRestart(); }

        bool IsRelevant(string id)
        {
            if (string.IsNullOrEmpty(id) || _cfg == null) return false;
            return string.Equals(id, _cfg.InputDeviceId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.MainDeviceId,  StringComparison.OrdinalIgnoreCase)
                || string.Equals(id, _cfg.AuxDeviceId,   StringComparison.OrdinalIgnoreCase);
        }

        void DebounceRestart() { _debounce.Stop(); _debounce.Start(); }
        void OnDebounceTick(object sender, EventArgs e) { _debounce.Stop(); StartOrRestart(); }

        void StartOrRestart()
        {
            Stop();
            if (_mm == null) _mm = new MMDeviceEnumerator();

            _inDev   = FindById(_cfg.InputDeviceId, DataFlow.Capture) ?? FindById(_cfg.InputDeviceId, DataFlow.Render);
            if (_inDev == null) _inDev = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _outMain = FindById(_cfg.MainDeviceId, DataFlow.Render);
            _outAux  = FindById(_cfg.AuxDeviceId,  DataFlow.Render);

            if (_outMain == null || _outAux == null)
            {
                MessageBox.Show("请先在“设置”里选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 输入设备初始化
            WaveFormat inFmt;
            if (_inDev.DataFlow == DataFlow.Capture)
            {
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                _capture = cap; inFmt = cap.WaveFormat;
            }
            else
            {
                var cap = new WasapiLoopbackCapture(_inDev);
                _capture = cap; inFmt = cap.WaveFormat;
            }

            // 初始化缓冲区
            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true };
            _bufAux  = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, ReadFully = true };

            // 创建主通道
            _mainOut = CreateOutWithPolicy(_outMain, AudioClientShareMode.Shared, _cfg.MainSync, 12, _bufMain, out _mainIsExclusive);

            // 创建副通道
            _auxOut = CreateOutWithPolicy(_outAux, AudioClientShareMode.Shared, _cfg.AuxSync, 180, _bufAux, out _auxIsExclusive);

            // 启动音频流
            _capture.StartRecording();
            _mainOut.Play();
            _auxOut.Play();

            _running = true;
        }

        void Stop()
        {
            if (!_running) return;
            try { _capture?.StopRecording(); } catch { }
            try { _mainOut?.Stop(); } catch { }
            try { _auxOut?.Stop();  } catch { }
            _running = false;
        }

        void DisposeAll()
        {
            try { _capture?.Dispose(); } catch { }
            try { _mainOut?.Dispose(); } catch { }
            try { _auxOut?.Dispose();  } catch { }
        }

        MMDevice FindById(string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            try
            {
                var col = _mm.EnumerateAudioEndPoints(flow, DeviceState.Active);
                foreach (var d in col) if (d.ID == id) return d;
            } catch { }
            return null;
        }

        WasapiOut CreateOutWithPolicy(MMDevice dev, AudioClientShareMode mode, SyncModeOption syncPref, int bufMs, IWaveProvider src, out bool eventUsed)
        {
            eventUsed = false;
            WasapiOut w;
            if (syncPref == SyncModeOption.Polling)
            {
                w = TryOut(dev, mode, false, bufMs, src);
                return w;
            }
            if (syncPref == SyncModeOption.Event)
            {
                w = TryOut(dev, mode, true, bufMs, src);
                if (w != null) { eventUsed = true; return w; }
                w = TryOut(dev, mode, false, bufMs, src);
                return w;
            }

            w = TryOut(dev, mode, true, bufMs, src);
            if (w != null) { eventUsed = true; return w; }
            return TryOut(dev, mode, false, bufMs, src);
        }

        WasapiOut TryOut(MMDevice dev, AudioClientShareMode mode, bool eventSync, int bufMs, IWaveProvider src)
        {
            try
            {
                var w = new WasapiOut(dev, mode, eventSync, bufMs);
                w.Init(src);
                return w;
            }
            catch
            {
                return null;
            }
        }

        public StatusSnapshot GetStatusSnapshot()
        {
            return new StatusSnapshot
            {
                Running = _running,
                InputRole = _inDev?.FriendlyName ?? "-",
                MainDevice = _outMain?.FriendlyName ?? "-",
                AuxDevice = _outAux?.FriendlyName ?? "-"
            };
        }

        public void Dispose()
        {
            Stop();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
            _debounce.Dispose();
        }
    }
}
