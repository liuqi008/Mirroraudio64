using System;
using System.Runtime.InteropServices;
using System.Threading;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace MirrorAudio.AppContextApp
{
    /// <summary>
    /// 采集输入（可选：环回），同时推送到两个输出（主/副），
    /// 支持独占与RAW直通；若"强制格式"=true则尝试以指定采样率&位深启动设备；否则按设备默认格式。
    /// 不做重采样：源与目标不匹配时，直接拒绝启动（以避免隐式重采样）。
    /// </summary>
    public sealed class AudioEngine : IDisposable
    {
        private readonly AppSettings _cfg;
        private readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        private IWaveIn? _capture;
        private IWaveProvider? _mainSource, _auxSource;
        private WasapiExclusivePlayer? _mainOut, _auxOut;
        private BufferedWaveProvider? _mainBuf, _auxBuf;

        public AudioEngine(AppSettings cfg) { _cfg = cfg; }

        public void Start()
        {
            // 1) 输入：优先显式选择；否则系统输出环回
            WaveFormat inputFormat;
            if (!string.IsNullOrEmpty(_cfg.InputDeviceId))
            {
                var inDev = _mm.GetDevice(_cfg.InputDeviceId);
                var cap = new WasapiCapture(inDev, false, 10);
                cap.ShareMode = AudioClientShareMode.Shared;
                cap.StartRecording();
                _capture = cap;
                inputFormat = cap.WaveFormat;
            }
            else
            {
                var defaultRender = _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                var cap = new WasapiLoopbackCapture(defaultRender);
                cap.StartRecording();
                _capture = cap;
                inputFormat = cap.WaveFormat;
            }

            // 2) 输出：主/副均创建独立缓冲；若禁用独占则仍用ExclusivePlayer包装（内部会选择共享）
            _mainBuf = new BufferedWaveProvider(inputFormat) { DiscardOnBufferOverflow = true };
            _auxBuf  = new BufferedWaveProvider(inputFormat) { DiscardOnBufferOverflow = true };

            _capture.DataAvailable += (_, e) =>
            {
                _mainBuf?.AddSamples(e.Buffer, 0, e.BytesRecorded);
                _auxBuf?.AddSamples(e.Buffer, 0, e.BytesRecorded);
            };

            _mainSource = _mainBuf;
            _auxSource  = _auxBuf;

            // 3) 启动输出设备（不重采样：要求输入与目标格式完全一致，或目标不强制格式）
            StartRenderer(ref _mainOut, _cfg.MainDeviceId, _cfg.MainExclusive, _cfg.MainRaw,
                          _cfg.MainForceFormat, _cfg.MainSampleRate, _cfg.MainBits, _cfg.MainBufferMs, _mainSource, "主通道");

            StartRenderer(ref _auxOut, _cfg.AuxDeviceId, _cfg.AuxExclusive, _cfg.AuxRaw,
                          _cfg.AuxForceFormat, _cfg.AuxSampleRate, _cfg.AuxBits, _cfg.AuxBufferMs, _auxSource, "副通道");
        }

        private void StartRenderer(ref WasapiExclusivePlayer? player,
                                   string? deviceId, bool exclusive, bool raw,
                                   bool forceFormat, int rate, int bits, int bufMs,
                                   IWaveProvider source, string tag)
        {
            var dev = !string.IsNullOrEmpty(deviceId)
                      ? _mm.GetDevice(deviceId)
                      : _mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

            // 如果强制格式，要求source格式完全一致；否则抛错，避免任何重采样。
            var fmt = source.WaveFormat;
            if (forceFormat)
            {
                if (!(fmt.SampleRate == rate && fmt.BitsPerSample == bits && fmt.Channels == 2))
                {
                    throw new InvalidOperationException($"{tag}：输入({fmt.SampleRate}/{fmt.BitsPerSample})与强制格式({rate}/{bits})不一致，为避免重采样已终止。可关闭“强制使用下方格式”或调整输入源。");
                }
            }

            // 独占期望格式（若不强制，则用输入格式）
            var desired = forceFormat ? WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, rate, 2, rate * 2 * (bits/8), 2 * (bits/8), bits)
                                      : fmt;

            player = new WasapiExclusivePlayer(dev, desired, exclusive, raw, bufMs, source);
            player.Start();
        }

        public void Dispose()
        {
            try
            {
                _mainOut?.Dispose();
                _auxOut?.Dispose();
                if (_capture is WasapiCapture wc) wc.StopRecording();
                if (_capture is WasapiLoopbackCapture wlc) wlc.StopRecording();
                _capture?.Dispose();
            }
            finally
            {
                _mainOut = _auxOut = null;
                _capture = null;
                _mainBuf = _auxBuf = null;
            }
        }
    }

    /// <summary>
    /// 轻量播放器：支持独占+RAW（IAudioClient3）+事件驱动，基于RenderClient手喂数据，不做重采样。
    /// </summary>
    internal sealed class WasapiExclusivePlayer : IDisposable
    {
        private readonly MMDevice _device;
        private readonly IWaveProvider _source;
        private readonly bool _exclusive;
        private readonly bool _raw;
        private readonly int _bufferMs;
        private readonly WaveFormat _format;

        private IAudioClient3? _client3;
        private IAudioRenderClient? _render;
        private EventWaitHandle? _event;
        private Thread? _thread;
        private bool _running;

        public WasapiExclusivePlayer(MMDevice device, WaveFormat format, bool exclusive, bool raw, int bufferMs, IWaveProvider source)
        {
            _device = device;
            _format = format;
            _exclusive = exclusive;
            _raw = raw;
            _bufferMs = Math.Max(2, bufferMs);
            _source = source;
        }

        public void Start()
        {
            // COM接口
            var clientObj = _device.AudioClient;
            _client3 = (IAudioClient3)clientObj.ComInterface;

            // RAW属性
            if (_raw)
            {
                var props = new AudioClientProperties
                {
                    cbSize = (uint)Marshal.SizeOf<AudioClientProperties>(),
                    bIsOffload = false,
                    eCategory = AudioStreamCategory.Other,
                    Options = AudioClientStreamOptions.Raw
                };
                _client3.SetClientProperties(ref props);
            }

            var shareMode = _exclusive ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared;
            var streamFlags = AudioClientStreamFlags.EventCallback;

            // 设置缓冲
            long hnsBuffer = (long)_bufferMs * 10000; // ms -> 100ns

            // 初始化（不做自动转换，格式必须设备可开）
            _client3.Initialize(
                shareMode,
                streamFlags,
                hnsBuffer,
                hnsBuffer,
                _format,
                Guid.Empty);

            // 事件 & 渲染客户端
            _event = new EventWaitHandle(false, EventResetMode.AutoReset);
            _client3.SetEventHandle(_event.SafeWaitHandle.DangerousGetHandle());

            _render = (IAudioRenderClient)clientObj.AudioRenderClient;

            // 计算帧大小
            int blockAlign = _format.BlockAlign;
            int bufferFrameCount = _client3.GetBufferSize();
            int bufferBytes = bufferFrameCount * blockAlign;

            _running = true;
            _thread = new Thread(() => Pump(bufferFrameCount, bufferBytes, blockAlign)) { IsBackground = true, Priority = ThreadPriority.Highest };
            _thread.Start();

            _client3.Start();
        }

        private void Pump(int bufferFrames, int bufferBytes, int blockAlign)
        {
            byte[] temp = new byte[bufferBytes];
            while (_running)
            {
                _event!.WaitOne(50);

                int padding = _client3!.GetCurrentPadding();
                int framesAvailable = bufferFrames - padding;
                if (framesAvailable <= 0) continue;

                int bytesNeeded = framesAvailable * blockAlign;
                int read = _source.Read(temp, 0, Math.Min(bytesNeeded, temp.Length));

                int framesToWrite = read / blockAlign;
                if (framesToWrite <= 0) continue;

                IntPtr pData = _render!.GetBuffer(framesToWrite);
                Marshal.Copy(temp, 0, pData, framesToWrite * blockAlign);
                _render.ReleaseBuffer(framesToWrite, AudioClientBufferFlags.None);
            }
        }

        public void Dispose()
        {
            try
            {
                _running = false;
                try { _client3?.Stop(); } catch { }
                if (_thread is not null && _thread.IsAlive) _thread.Join(200);
            }
            finally
            {
                _render = null;
                _client3 = null;
                _event?.Dispose();
                _event = null;
            }
        }
    }

    #region CoreAudio interop (IAudioClient3 / RAW)

    [ComImport]
    [Guid("7ED4EE07-8E67-4CD4-BE3C-DBFE46A8D338")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioClient3
    {
        // IAudioClient methods
        int Initialize(AudioClientShareMode ShareMode, AudioClientStreamFlags StreamFlags, long hnsBufferDuration, long hnsPeriodicity, [In] WaveFormat waveFormat, Guid AudioSessionGuid);
        int GetBufferSize();
        int GetStreamLatency(out long phnsLatency);
        int GetCurrentPadding();
        int IsFormatSupported(AudioClientShareMode ShareMode, [In] WaveFormat pFormat, IntPtr ppClosestMatch);
        int GetMixFormat(out IntPtr ppDeviceFormat);
        int GetDevicePeriod(out long phnsDefaultDevicePeriod, out long phnsMinimumDevicePeriod);
        int Start();
        int Stop();
        int Reset();
        int SetEventHandle(IntPtr eventHandle);
        int GetService(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        // IAudioClient2/3
        int SetClientProperties(ref AudioClientProperties pProperties);
        int GetSharedModeEnginePeriod([In] WaveFormat pFormat, out int pDefaultPeriodInFrames, out int pFundamentalPeriodInFrames, out int pMinPeriodInFrames, out int pMaxPeriodInFrames);
        int GetSharedModeEnginePeriod(out IntPtr ppFormat, out int pDefaultPeriodInFrames, out int pFundamentalPeriodInFrames, out int pMinPeriodInFrames, out int pMaxPeriodInFrames);
        int InitializeSharedAudioStream(AudioClientStreamFlags StreamFlags, int PeriodInFrames, [In] WaveFormat pFormat, Guid AudioSessionGuid);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct AudioClientProperties
    {
        public uint cbSize;
        [MarshalAs(UnmanagedType.Bool)] public bool bIsOffload;
        public AudioStreamCategory eCategory;
        public AudioClientStreamOptions Options;
    }

    internal enum AudioClientStreamOptions
    {
        None = 0x0,
        Raw = 0x1,
        MatchFormat = 0x2
    }

    internal enum AudioStreamCategory
    {
        Other = 0,
    }

    [Flags]
    internal enum AudioClientStreamFlags : int
    {
        None = 0x0,
        EventCallback = 0x00040000,
    }

    [ComImport]
    [Guid("F294ACFC-3146-4483-A7BF-ADDCA7C260E2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioRenderClient
    {
        IntPtr GetBuffer(int numFramesRequested);
        void ReleaseBuffer(int numFramesWritten, AudioClientBufferFlags dwFlags);
    }

    [Flags]
    internal enum AudioClientBufferFlags
    {
        None = 0x0
    }

    #endregion
}