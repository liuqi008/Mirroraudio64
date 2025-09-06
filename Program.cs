
using System;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace MirrorAudio
{
    [Serializable]
    public enum ShareModeOption { Auto, Exclusive, Shared }
    [Serializable]
    public enum SyncModeOption  { Auto, Event, Polling }

    public sealed class AppSettings
    {
        public string MainDeviceId = null;
        public string AuxDeviceId  = null;
        public ShareModeOption MainShare = ShareModeOption.Exclusive;
        public ShareModeOption AuxShare  = ShareModeOption.Exclusive;
        public SyncModeOption  MainSync  = SyncModeOption.Event;
        public SyncModeOption  AuxSync   = SyncModeOption.Event;
        public int MainBufMs = 6;
        public int AuxBufMs  = 6;
        public int MainRate  = 192000;
        public int AuxRate   = 44100;
        public int MainBits  = 24;
        public int AuxBits   = 16;
        public int Channels  = 2;
    }

    public sealed class StatusSnapshot
    {
        public bool Running;
        public string MainDevice, AuxDevice;
        public string MainMode,   AuxMode;
        public string MainSync,   AuxSync;
        public string MainFormat, AuxFormat;
        public int MainBufferMs,  AuxBufferMs;
        public double MainDefaultPeriodMs, MainMinimumPeriodMs;
        public double AuxDefaultPeriodMs,  AuxMinimumPeriodMs;
        public bool MainRaw, AuxRaw;
    }

    internal sealed class TrayApp : IDisposable
    {
        readonly NotifyIcon _ni;
        readonly ContextMenuStrip _menu;
        readonly ToolStripMenuItem _miStart, _miStop, _miSettings, _miExit;

        readonly MMDeviceEnumerator _devEnum = new MMDeviceEnumerator();
        WasapiOut _mainOut, _auxOut;
        IWaveProvider _mainSrc, _auxSrc;
        bool _running, _mainIsExclusive, _auxIsExclusive;
        bool _mainEventSyncUsed, _auxEventSyncUsed;
        bool _mainRawEnabled, _auxRawEnabled;
        int  _mainBufEffectiveMs, _auxBufEffectiveMs;
        string _mainFmtStr, _auxFmtStr;
        MMDevice _outMain, _outAux;
        double _defMainMs, _minMainMs, _defAuxMs, _minAuxMs;

        public AppSettings Settings = new AppSettings();

        public TrayApp()
        {
            _menu = new ContextMenuStrip();
            _miStart    = new ToolStripMenuItem("开始(&S)", null, (s,e)=>StartOrRestart());
            _miStop     = new ToolStripMenuItem("停止(&T)", null, (s,e)=>Stop());
            _miSettings = new ToolStripMenuItem("设置(&G)", null, (s,e)=>ShowSettings());
            _miExit     = new ToolStripMenuItem("退出(&X)", null, (s,e)=>{ Stop(); Application.Exit(); });
            _menu.Items.AddRange(new ToolStripItem[]{ _miStart, _miStop, new ToolStripSeparator(), _miSettings, new ToolStripSeparator(), _miExit });
            _ni = new NotifyIcon { Text = "MirrorAudio", Visible = true, ContextMenuStrip = _menu };
            try { _ni.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch {}
        }

        public void Dispose()
        {
            Stop();
            if (_ni != null) { _ni.Visible = false; _ni.Dispose(); }
            if (_devEnum != null) _devEnum.Dispose();
        }

        void ShowSettings()
        {
            using (var dlg = new SettingsForm(Settings, GetStatusSnapshot))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    Settings = dlg.Result ?? Settings;
                    StartOrRestart();
                }
            }
        }

        public void StartOrRestart()
        {
            Stop();
            _mainRawEnabled = _auxRawEnabled = false;
            _mainIsExclusive = _auxIsExclusive = false;
            _mainEventSyncUsed = _auxEventSyncUsed = false;
            _mainBufEffectiveMs = _auxBufEffectiveMs = 0;
            _mainFmtStr = _auxFmtStr = "-";
            _defMainMs = _minMainMs = _defAuxMs = _minAuxMs = 0;

            try
            {
                _outMain = PickRenderDevice(Settings.MainDeviceId);
                _outAux  = PickRenderDevice(Settings.AuxDeviceId);

                QueryPeriods(_outMain, out _defMainMs, out _minMainMs);
                QueryPeriods(_outAux,  out _defAuxMs,  out _minAuxMs);

                _mainSrc = BuildSineSource(Settings.MainRate, Settings.MainBits, Settings.Channels, 220.0);
                _auxSrc  = BuildSineSource(Settings.AuxRate,  Settings.AuxBits,  Settings.Channels, 440.0);

                _mainOut = BuildWasapiWithRawFallback(_outMain, Settings.MainShare, Settings.MainSync, Settings.MainBufMs, _mainSrc,
                                                      out _mainIsExclusive, out _mainEventSyncUsed, out _mainRawEnabled,
                                                      out _mainBufEffectiveMs, out _mainFmtStr);

                _auxOut = BuildWasapiWithRawFallback(_outAux, Settings.AuxShare, Settings.AuxSync, Settings.AuxBufMs, _auxSrc,
                                                     out _auxIsExclusive, out _auxEventSyncUsed, out _auxRawEnabled,
                                                     out _auxBufEffectiveMs, out _auxFmtStr);

                if (_mainOut != null) _mainOut.Play();
                if (_auxOut  != null) _auxOut.Play();
                _running = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败: " + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Stop();
            }
        }

        public void Stop()
        {
            _running = false;
            if (_mainOut != null) { try { _mainOut.Stop(); } catch {} _mainOut.Dispose(); _mainOut = null; }
            if (_auxOut  != null) { try { _auxOut.Stop(); }  catch {} _auxOut.Dispose();  _auxOut  = null; }
        }

        MMDevice PickRenderDevice(string idOrNull)
        {
            if (!string.IsNullOrEmpty(idOrNull))
            {
                try { return _devEnum.GetDevice(idOrNull); } catch {}
            }
            return _devEnum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        }

        void QueryPeriods(MMDevice dev, out double defMs, out double minMs)
        {
            defMs = 0; minMs = 0;
            if (dev == null) return;
            try
            {
                var ac = dev.AudioClient;
                var t = ac.GetType();

                var pDef = t.GetProperty("DefaultDevicePeriod");
                var pMin = t.GetProperty("MinimumDevicePeriod");
                if (pDef != null && pMin != null)
                {
                    long defTicks = Convert.ToInt64(pDef.GetValue(ac, null));
                    long minTicks = Convert.ToInt64(pMin.GetValue(ac, null));
                    defMs = defTicks / 10000.0;
                    minMs = minTicks / 10000.0;
                    return;
                }

                var mGet = t.GetMethod("GetDevicePeriod");
                if (mGet != null)
                {
                    object[] args = new object[] { 0L, 0L };
                    mGet.Invoke(ac, args);
                    long defTicks = Convert.ToInt64(args[0]);
                    long minTicks = Convert.ToInt64(args[1]);
                    defMs = defTicks / 10000.0;
                    minMs = minTicks / 10000.0;
                    return;
                }

                var mDef = t.GetMethod("GetDefaultDevicePeriod");
                var mMin = t.GetMethod("GetMinimumDevicePeriod");
                if (mDef != null && mMin != null)
                {
                    object[] a1 = new object[] { 0L };
                    object[] a2 = new object[] { 0L };
                    mDef.Invoke(ac, a1);
                    mMin.Invoke(ac, a2);
                    long defTicks = Convert.ToInt64(a1[0]);
                    long minTicks = Convert.ToInt64(a2[0]);
                    defMs = defTicks / 10000.0;
                    minMs = minTicks / 10000.0;
                    return;
                }
            }
            catch
            {
                // leave 0
            }
        }

        IWaveProvider BuildSineSource(int rate, int bits, int ch, double freq)
        {
            var wf = new WaveFormat(rate, bits, ch);
            return new SignalGeneratorProvider(wf, freq);
        }

        WasapiOut BuildWasapiWithRawFallback(MMDevice dev, ShareModeOption share, SyncModeOption syncPref, int reqBufMs, IWaveProvider src,
                                             out bool isExclusive, out bool eventUsed, out bool rawEnabled,
                                             out int effBufMs, out string fmtStr)
        {
            isExclusive = false; eventUsed = false; rawEnabled = false; effBufMs = 0; fmtStr = "-";
            if (dev == null || src == null) return null;

            var tryExclusive = (share != ShareModeOption.Shared);
            var allowSharedFallback = (share != ShareModeOption.Exclusive);

            WasapiOut outp = null;

            if (tryExclusive)
            {
                if (TrySetRaw(dev))
                {
                    outp = CreateOut(dev, AudioClientShareMode.Exclusive, syncPref, reqBufMs, src, out eventUsed, out effBufMs, out fmtStr);
                    if (outp != null) { isExclusive = true; rawEnabled = true; return outp; }
                    ClearRaw(dev);
                }

                outp = CreateOut(dev, AudioClientShareMode.Exclusive, syncPref, reqBufMs, src, out eventUsed, out effBufMs, out fmtStr);
                if (outp != null) { isExclusive = true; return outp; }
            }

            if (allowSharedFallback)
            {
                outp = CreateOut(dev, AudioClientShareMode.Shared, syncPref, reqBufMs, src, out eventUsed, out effBufMs, out fmtStr);
                if (outp != null) { isExclusive = false; return outp; }
            }

            return null;
        }

        WasapiOut CreateOut(MMDevice dev, AudioClientShareMode mode, SyncModeOption pref, int reqBufMs, IWaveProvider src,
                            out bool eventUsed, out int effBufMs, out string fmtStr)
        {
            eventUsed = false; effBufMs = 0; fmtStr = "-";
            WasapiOut w = null;
            try
            {
                bool useEvent = (pref != SyncModeOption.Polling);
                var wo = new WasapiOut(dev, mode, useEvent, reqBufMs);
                wo.Init(src);
                w = wo;
                effBufMs = reqBufMs;
                eventUsed = useEvent;
                fmtStr = FormatStringFromWaveProvider(src);
            }
            catch
            {
                if (w != null) { try { w.Dispose(); } catch {} }
                w = null;
            }
            return w;
        }

        string FormatStringFromWaveProvider(IWaveProvider wp)
        {
            try
            {
                var wf = wp.WaveFormat;
                return string.Format("{0}ch {1}bit {2}Hz", wf.Channels, wf.BitsPerSample, wf.SampleRate);
            }
            catch { return "-"; }
        }

        bool TrySetRaw(MMDevice dev)
        {
            try
            {
                var ac = dev.AudioClient;
                var acType = ac.GetType();
                var propsType = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientProperties");
                var catType   = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientStreamCategory");
                var optType   = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientStreamOptions");
                var setMethod = acType.GetMethod("SetClientProperties");
                if (propsType == null || catType == null || optType == null || setMethod == null) return false;

                var props = Activator.CreateInstance(propsType);
                var catMedia = Enum.Parse(catType, "Media");
                propsType.GetProperty("Category").SetValue(props, catMedia, null);
                var optRaw = Enum.Parse(optType, "Raw");
                propsType.GetProperty("Options").SetValue(props, optRaw, null);
                setMethod.Invoke(ac, new object[] { props });
                return true;
            }
            catch { return false; }
        }

        void ClearRaw(MMDevice dev)
        {
            try
            {
                var ac = dev.AudioClient;
                var acType = ac.GetType();
                var propsType = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientProperties");
                var catType   = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientStreamCategory");
                var optType   = acType.Assembly.GetType("NAudio.CoreAudioApi.AudioClientStreamOptions");
                var setMethod = acType.GetMethod("SetClientProperties");
                if (propsType == null || catType == null || optType == null || setMethod == null) return;

                var props = Activator.CreateInstance(propsType);
                var catMedia = Enum.Parse(catType, "Media");
                var optNone  = Enum.Parse(optType, "None");
                propsType.GetProperty("Category").SetValue(props, catMedia, null);
                propsType.GetProperty("Options").SetValue(props, optNone,  null);
                setMethod.Invoke(ac, new object[] { props });
            }
            catch {}
        }

        public StatusSnapshot GetStatusSnapshot()
        {
            return new StatusSnapshot
            {
                Running = _running,
                MainDevice = SafeName(_outMain),
                AuxDevice  = SafeName(_outAux),
                MainMode = _mainOut!=null? (_mainIsExclusive? "独占":"共享") : "-",
                AuxMode  = _auxOut !=null? (_auxIsExclusive ? "独占":"共享") : "-",
                MainSync = _mainOut!=null? (_mainEventSyncUsed? "事件":"轮询") : "-",
                AuxSync  = _auxOut !=null? (_auxEventSyncUsed ? "事件":"轮询") : "-",
                MainFormat = _mainOut!=null? _mainFmtStr : "-",
                AuxFormat  = _auxOut !=null? _auxFmtStr  : "-",
                MainBufferMs = _mainOut!=null? _mainBufEffectiveMs : 0,
                AuxBufferMs  = _auxOut !=null? _auxBufEffectiveMs  : 0,
                MainDefaultPeriodMs = _defMainMs,
                MainMinimumPeriodMs = _minMainMs,
                AuxDefaultPeriodMs  = _defAuxMs,
                AuxMinimumPeriodMs  = _minAuxMs,
                MainRaw = _mainRawEnabled,
                AuxRaw  = _auxRawEnabled
            };
        }

        string SafeName(MMDevice d)
        {
            if (d == null) return "-";
            try { return d.FriendlyName; } catch { return "(unknown)"; }
        }
    }

    internal sealed class SignalGeneratorProvider : IWaveProvider
    {
        readonly WaveFormat _format;
        readonly double _freq;
        double _t;

        public SignalGeneratorProvider(WaveFormat format, double freq)
        {
            _format = format;
            _freq = freq;
        }

        public WaveFormat WaveFormat { get { return _format; } }

        public int Read(byte[] buffer, int offset, int count)
        {
            int bytesPerSample = _format.BitsPerSample / 8;
            int samples = count / bytesPerSample;
            int ch = _format.Channels;
            int frames = samples / ch;
            double sr = _format.SampleRate;

            for (int i = 0; i < frames; i++)
            {
                double sample = Math.Sin(2.0 * Math.PI * _freq * _t / sr);
                for (int c = 0; c < ch; c++)
                {
                    WriteSample(buffer, offset + (i * ch + c) * bytesPerSample, sample, bytesPerSample);
                }
                _t += 1.0;
            }
            return frames * ch * bytesPerSample;
        }

        void WriteSample(byte[] buf, int ofs, double val, int bps)
        {
            if (val > 1) val = 1;
            if (val < -1) val = -1;

            if (bps == 2)
            {
                short s = (short)(val * short.MaxValue);
                buf[ofs+0] = (byte)(s & 0xFF);
                buf[ofs+1] = (byte)((s >> 8) & 0xFF);
            }
            else if (bps == 3)
            {
                int s = (int)(val * 8388607.0);
                buf[ofs+0] = (byte)(s & 0xFF);
                buf[ofs+1] = (byte)((s >> 8) & 0xFF);
                buf[ofs+2] = (byte)((s >> 16) & 0xFF);
            }
            else
            {
                int s = (int)(val * int.MaxValue);
                buf[ofs+0] = (byte)(s & 0xFF);
                buf[ofs+1] = (byte)((s >> 8) & 0xFF);
                buf[ofs+2] = (byte)((s >> 16) & 0xFF);
                buf[ofs+3] = (byte)((s >> 24) & 0xFF);
            }
        }
    }

    static class Program
    {
        static Mutex _mtx;
        [STAThread]
        static void Main()
        {
            bool ok;
            _mtx = new Mutex(true, "Global\\MirrorAudio_{7D21A2D9-6C1D-4C2A-9A49-6F9D3092B3F7}", out ok);
            if (!ok) return;

            try { MediaFoundationApi.Startup(); } catch {}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (var app = new TrayApp())
            {
                Application.Run();
            }

            try { MediaFoundationApi.Shutdown(); } catch {}
        }
    }
}
