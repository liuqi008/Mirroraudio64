using System;

namespace MirrorAudio
{
    public enum ShareModeOption { Auto = 0, Exclusive = 1, Shared = 2 }
    public enum SyncModeOption { Auto = 0, Event = 1, Polling = 2 }
    public enum BufferAlignMode { DefaultAlign = 0, MinAlign = 1 }
    public enum InputFormatStrategy { SystemMix = 0, Custom = 1, Float32Prefer = 2 }

    public class AppSettings
    {
        // Device Ids (optional, not used by current code paths but kept for compatibility)
        public string InputDeviceId;
        public string MainDeviceId;
        public string AuxDeviceId;

        // Output modes
        public ShareModeOption MainShare = ShareModeOption.Auto;
        public ShareModeOption AuxShare = ShareModeOption.Auto;
        public SyncModeOption MainSync = SyncModeOption.Auto;
        public SyncModeOption AuxSync = SyncModeOption.Auto;

        // Formats
        public int MainRate = 48000;
        public int MainBits = 24;
        public int AuxRate = 48000;
        public int AuxBits = 24;

        // Buffers
        public int MainBufMs = 12;
        public int AuxBufMs = 150;
        public BufferAlignMode MainBufMode = BufferAlignMode.MinAlign;
        public BufferAlignMode AuxBufMode = BufferAlignMode.DefaultAlign;

        // Resampler controls
        public int MainResamplerQuality = 50; // 60/50/40/30 in UI
        public int AuxResamplerQuality = 30;
        public bool MainForceInternalResamplerInShared = false;
        public bool AuxForceInternalResamplerInShared = false;

        // Misc
        public bool AutoStart = false;
        public bool EnableLogging = true;

        // Input strategy
        public InputFormatStrategy InputFormatStrategy = InputFormatStrategy.SystemMix;
        public int InputCustomSampleRate = 48000;
        public int InputCustomBitDepth = 24;

        public AppSettings Clone() => (AppSettings)this.MemberwiseClone();
    }

    public class StatusSnapshot
    {
        public bool MainInternalResampler { get; set; }
        public bool AuxInternalResampler  { get; set; }
        public bool MainMultiSRC { get; set; }
        public bool AuxMultiSRC  { get; set; }
    }
}