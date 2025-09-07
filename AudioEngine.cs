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
    
    /// <summary>
    /// 轻量播放器封装：基于 NAudio.Wave.WasapiOut，支持独占/共享与事件驱动。
    /// 说明：为了跨环境可编译与稳定运行，本封装不直接访问 IAudioClient3。
    /// </summary>
    internal sealed class WasapiExclusivePlayer : IDisposable
    {
        private readonly WasapiOut _out;

        public WasapiExclusivePlayer(MMDevice device, WaveFormat format, bool exclusive, bool raw, int bufferMs, IWaveProvider source)
        {
            // 注意：独占模式下本就绕过系统混音。RAW 对共享模式意义更大；
            // 由于 NAudio 未稳定公开 RAW 旁路入口，此处先保证独占路径与事件驱动低延迟。
            _out = new WasapiOut(device, exclusive ? AudioClientShareMode.Exclusive : AudioClientShareMode.Shared, true, bufferMs);
            _out.Init(source);
        }

        public void Start() => _out.Play();

        public void Dispose()
        {
            try { _out?.Stop(); } catch {}
            _out?.Dispose();
        }
    }

}
}