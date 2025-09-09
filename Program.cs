using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace MirrorAudio
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp())
            {
                Application.Run();
            }
        }
    }

    public class TrayApp : IDisposable
    {
        private NotifyIcon _icon;
        private ContextMenuStrip _menu;

        private AppSettings _cfg = new AppSettings();
        private bool _running;

        // 输入
        private MMDevice _inDev;
        private WasapiCapture _capture;

        // 主输出
        private MMDevice _outMain;
        private WasapiOut _mainOut;
        private BufferedWaveProvider _bufMain;
        private MediaFoundationResampler _resMain;
        private bool _mainEventSyncUsed;
        private bool _mainResampling;
        private bool _mainNoSRC;
        private ShareMode _mainShareModeNA; // NAudio
        private AudioClientShareMode _mainShareMode; // 记录是独占/共享

        // 副输出
        private MMDevice _outAux;
        private WasapiOut _auxOut;
        private BufferedWaveProvider _bufAux;
        private MediaFoundationResampler _resAux;
        private bool _auxEventSyncUsed;
        private bool _auxResampling;
        private bool _auxNoSRC;
        private ShareMode _auxShareModeNA;
        private AudioClientShareMode _auxShareMode;

        // 设备管理
        private MMDeviceEnumerator _enum = new MMDeviceEnumerator();

        public TrayApp()
        {
            BuildTray();
            StartOrRestart();
        }

        private void BuildTray()
        {
            _menu = new ContextMenuStrip();
            var mOpen = new ToolStripMenuItem("打开设置", null, (_, __) => OpenSettings());
            var mStart = new ToolStripMenuItem("启动/重启", null, (_, __) => StartOrRestart());
            var mStop = new ToolStripMenuItem("停止", null, (_, __) => Stop());
            var mExit = new ToolStripMenuItem("退出", null, (_, __) => { Stop(); Application.Exit(); });

            _menu.Items.AddRange(new ToolStripItem[] { mOpen, mStart, mStop, new ToolStripSeparator(), mExit });

            _icon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "MirrorAudio",
                ContextMenuStrip = _menu
            };
        }

        private void OpenSettings()
        {
            using (var f = new SettingsForm(_cfg, GetStatusSnapshot))
            {
                f.ApplyRequested += (sNew) =>
                {
                    _cfg = sNew;
                    // TODO: 保存配置、处理自启动
                    StartOrRestart();
                };

                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result;
                    // TODO: 保存配置、处理自启动
                    StartOrRestart();
                }
            }
        }

        private void Stop()
        {
            _running = false;

            try { _capture?.StopRecording(); } catch { }
            _capture?.Dispose(); _capture = null;

            _mainOut?.Stop(); _mainOut?.Dispose(); _mainOut = null;
            _auxOut?.Stop(); _auxOut?.Dispose(); _auxOut = null;

            _resMain?.Dispose(); _resMain = null;
            _resAux?.Dispose(); _resAux = null;

            _bufMain = null; _bufAux = null;
        }

        public void StartOrRestart()
        {
            Stop();
            try
            {
                // —— 设备选择（简化：取默认）——
                _inDev = _enum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia); // 环回抓系统混音
                _outMain = _enum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                _outAux  = _enum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

                // —— 输入：固定共享（环回）——
                _capture = new WasapiLoopbackCapture(_inDev);
                if (_cfg.InputFormatStrategy == InputFormatStrategy.Custom)
                {
                    // 尝试设置输入格式（并非所有驱动都接受）
                    _capture.WaveFormat = new WaveFormat(_cfg.InputCustomSampleRate, _cfg.InputCustomBitDepth, 2);
                }
                _capture.DataAvailable += Capture_DataAvailable;
                _capture.RecordingStopped += (_, __) => { };
                _capture.StartRecording();

                // —— 主/副缓冲（写入输入的数据）——
                _bufMain = new BufferedWaveProvider(_capture.WaveFormat) { DiscardOnBufferOverflow = true, BufferLength = 1 << 18 };
                _bufAux  = new BufferedWaveProvider(_capture.WaveFormat) { DiscardOnBufferOverflow = true, BufferLength = 1 << 18 };

                // —— 主输出 —— 
                BuildMainOutput();

                // —— 副输出 —— 
                BuildAuxOutput();

                _running = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败：\n" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        private void BuildMainOutput()
        {
            _resMain?.Dispose(); _resMain = null;
            _mainOut?.Dispose(); _mainOut = null;

            var share = _cfg.MainShare == ShareModeOption.Exclusive ? AudioClientShareMode.Exclusive :
                        _cfg.MainShare == ShareModeOption.Shared   ? AudioClientShareMode.Shared : AudioClientShareMode.Shared;
            _mainShareMode = share;

            var ms = BufAligned(_cfg.MainBufMs, _cfg.MainBufMode);
            IWaveProvider feed = _bufMain;

            if (share == AudioClientShareMode.Shared)
            {
                if (_cfg.MainForceInternalResamplerInShared)
                {
                    var target = TryGetMixFormat(_outMain) ?? new WaveFormat(48000, _cfg.MainBits, 2);
                    if (!WaveFormatEquals(feed.WaveFormat, target))
                    {
                        _resMain = new MediaFoundationResampler(feed, target) { ResamplerQuality = _cfg.MainResamplerQuality };
                        feed = _resMain;
                        _mainResampling = true; _mainNoSRC = false;
                    }
                }
                _mainOut = new WasapiOut(_outMain, ShareMode.Shared, _cfg.MainSync == SyncModeOption.Event, ms);
                _mainOut.Init(feed);
                _mainOut.Play();
            }
            else // Exclusive
            {
                var desired = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
                if (!WaveFormatEquals(feed.WaveFormat, desired))
                {
                    _resMain = new MediaFoundationResampler(feed, desired) { ResamplerQuality = _cfg.MainResamplerQuality };
                    feed = _resMain;
                    _mainResampling = true; _mainNoSRC = false;
                }
                _mainOut = new WasapiOut(_outMain, ShareMode.Exclusive, _cfg.MainSync == SyncModeOption.Event, ms);
                _mainOut.Init(feed);
                _mainOut.Play();
            }
        }

        private void BuildAuxOutput()
        {
            _resAux?.Dispose(); _resAux = null;
            _auxOut?.Dispose(); _auxOut = null;

            var share = _cfg.AuxShare == ShareModeOption.Exclusive ? AudioClientShareMode.Exclusive :
                        _cfg.AuxShare == ShareModeOption.Shared   ? AudioClientShareMode.Shared : AudioClientShareMode.Shared;
            _auxShareMode = share;

            var ms = BufAligned(_cfg.AuxBufMs, _cfg.AuxBufMode);
            IWaveProvider feed = _bufAux;

            if (share == AudioClientShareMode.Shared)
            {
                if (_cfg.AuxForceInternalResamplerInShared)
                {
                    var target = TryGetMixFormat(_outAux) ?? new WaveFormat(48000, _cfg.AuxBits, 2);
                    if (!WaveFormatEquals(feed.WaveFormat, target))
                    {
                        _resAux = new MediaFoundationResampler(feed, target) { ResamplerQuality = _cfg.AuxResamplerQuality };
                        feed = _resAux;
                        _auxResampling = true; _auxNoSRC = false;
                    }
                }
                _auxOut = new WasapiOut(_outAux, ShareMode.Shared, _cfg.AuxSync == SyncModeOption.Event, ms);
                _auxOut.Init(feed);
                _auxOut.Play();
            }
            else // Exclusive
            {
                var desired = new WaveFormat(_cfg.AuxRate, _cfg.AuxBits, 2);
                if (!WaveFormatEquals(feed.WaveFormat, desired))
                {
                    _resAux = new MediaFoundationResampler(feed, desired) { ResamplerQuality = _cfg.AuxResamplerQuality };
                    feed = _resAux;
                    _auxResampling = true; _auxNoSRC = false;
                }
                _auxOut = new WasapiOut(_outAux, ShareMode.Exclusive, _cfg.AuxSync == SyncModeOption.Event, ms);
                _auxOut.Init(feed);
                _auxOut.Play();
            }
        }

        private WaveFormat TryGetMixFormat(MMDevice dev)
        {
            try { return dev.AudioClient.MixFormat; } catch { return null; }
        }

        private void Capture_DataAvailable(object sender, WaveInEventArgs e)
        {
            // 同步写入主/副环形缓冲；不足则自然丢弃（DiscardOnBufferOverflow=true）
            _bufMain?.AddSamples(e.Buffer, 0, e.BytesRecorded);
            _bufAux ?.AddSamples(e.Buffer, 0, e.BytesRecorded);
        }

        private int BufAligned(int reqMs, BufferAlignMode mode)
        {
            // 简化：不查询设备最小周期，直接限幅
            var v = Math.Max(3, Math.Min(500, reqMs));
            if (mode == BufferAlignMode.MinAlign)
            {
                // 近似对齐到 3/4 倍（你项目里可改为“设备最小周期×3/4”精确对齐）
                if (v < 12) v = 12;
            }
            return v;
        }

        public StatusSnapshot GetStatusSnapshot()
        {
            // 多次 SRC：仅在共享且我们内部已做过重采样，且目标 != 设备 MixFormat
            bool mainInternal = (_resMain != null);
            bool auxInternal  = (_resAux  != null);

            bool mainMulti = false;
            if (_mainOut != null && _mainShareMode == AudioClientShareMode.Shared && mainInternal)
            {
                var mix = TryGetMixFormat(_outMain);
                var final = _resMain?.OutputWaveFormat ?? _bufMain?.WaveFormat;
                mainMulti = (mix != null) && !WaveFormatEquals(final, mix);
            }

            bool auxMulti = false;
            if (_auxOut != null && _auxShareMode == AudioClientShareMode.Shared && auxInternal)
            {
                var mix = TryGetMixFormat(_outAux);
                var final = _resAux?.OutputWaveFormat ?? _bufAux?.WaveFormat;
                auxMulti = (mix != null) && !WaveFormatEquals(final, mix);
            }

            return new StatusSnapshot
            {
                MainInternalResampler = mainInternal,
                AuxInternalResampler  = auxInternal,
                MainMultiSRC = mainMulti,
                AuxMultiSRC  = auxMulti
            };
        }

        private static bool WaveFormatEquals(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate
                && a.Channels == b.Channels
                && a.BitsPerSample == b.BitsPerSample;
        }

        public void Dispose()
        {
            Stop();
            _icon?.Dispose();
            _enum?.Dispose();
        }
    }

    // —— 与 SettingsForm.cs 保持一致的模型枚举 —— //
    public enum ShareModeOption { Auto = 0, Exclusive = 1, Shared = 2 }
    public enum SyncModeOption { Auto = 0, Event = 1, Polling = 2 }
    public enum BufferAlignMode { DefaultAlign = 0, MinAlign = 1 }
    public enum InputFormatStrategy { SystemMix = 0, Custom = 1, Float32Prefer = 2 }

    public class AppSettings
    {
        public string InputDeviceId;
        public string MainDeviceId;
        public string AuxDeviceId;

        public ShareModeOption MainShare = ShareModeOption.Auto;
        public ShareModeOption AuxShare = ShareModeOption.Auto;
        public SyncModeOption MainSync = SyncModeOption.Auto;
        public SyncModeOption AuxSync = SyncModeOption.Auto;

        public int MainRate = 48000;
        public int MainBits = 24;
        public int AuxRate = 48000;
        public int AuxBits = 24;

        public int MainBufMs = 12;
        public int AuxBufMs = 150;
        public BufferAlignMode MainBufMode = BufferAlignMode.MinAlign;
        public BufferAlignMode AuxBufMode = BufferAlignMode.DefaultAlign;

        public int MainResamplerQuality = 50;
        public int AuxResamplerQuality = 30;
        public bool MainForceInternalResamplerInShared = false;
        public bool AuxForceInternalResamplerInShared = false;

        public bool AutoStart = false;
        public bool EnableLogging = true;

        public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        public int InputCustomSampleRate = 48000;
        public int InputCustomBitDepth = 24;

        public AppSettings Clone() => (AppSettings)MemberwiseClone();
    }

    public class StatusSnapshot
    {
        public bool MainInternalResampler { get; set; }
        public bool AuxInternalResampler { get; set; }
        public bool MainMultiSRC { get; set; }
        public bool AuxMultiSRC { get; set; }
    }
}
