using System;
using System.Drawing;
using System.Windows.Forms;

namespace MirrorAudio
{
    public partial class SettingsForm : Form
    {
        private readonly Func<StatusSnapshot> _statusProvider;
        private AppSettings _settings;

        // --- Left status ---
        private Label lblTitleMain;
        private Label lblMainLine1;
        private Label lblMainLine2; // 质量
        private Label lblTitleAux;
        private Label lblAuxLine1;
        private Label lblAuxLine2;  // 质量

        // --- Right settings ---
        private GroupBox grpMain;
        private ComboBox cboMainRate;
        private ComboBox cboMainBits;
        private NumericUpDown nudMainBuf;
        private ComboBox cboMainShare;
        private ComboBox cboMainSync;
        private ComboBox cboMainAlign;
        private ComboBox cboMainQuality;
        private CheckBox chkMainForceSharedResampler;

        private GroupBox grpAux;
        private ComboBox cboAuxRate;
        private ComboBox cboAuxBits;
        private NumericUpDown nudAuxBuf;
        private ComboBox cboAuxShare;
        private ComboBox cboAuxSync;
        private ComboBox cboAuxAlign;
        private ComboBox cboAuxQuality;
        private CheckBox chkAuxForceSharedResampler;

        private Button btnRefresh;
        private Button btnOk;
        private Button btnCancel;

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings current, Func<StatusSnapshot> statusProvider)
        {
            _settings = current;
            _statusProvider = statusProvider;
            InitializeComponent();
            LoadFromSettings();
            RenderStatus();
        }

        private void InitializeComponent()
        {
            this.Text = "MirrorAudio 设置";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ClientSize = new Size(860, 520);

            // Left panel (status)
            var left = new Panel { Left = 10, Top = 10, Width = 400, Height = 500, BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(left);

            lblTitleMain = new Label { Left = 10, Top = 10, Width = 360, Text = "主通道状态", Font = new Font(Font, FontStyle.Bold) };
            lblMainLine1 = new Label { Left = 10, Top = 35, Width = 360, Text = "主：直通/重采样/模式/格式/缓冲" };
            lblMainLine2 = new Label { Left = 10, Top = 55, Width = 360, Text = "主：程序内重采样质量 / 多次SRC" };

            lblTitleAux = new Label { Left = 10, Top = 100, Width = 360, Text = "副通道状态", Font = new Font(Font, FontStyle.Bold) };
            lblAuxLine1  = new Label { Left = 10, Top = 125, Width = 360, Text = "副：直通/重采样/模式/格式/缓冲" };
            lblAuxLine2  = new Label { Left = 10, Top = 145, Width = 360, Text = "副：程序内重采样质量 / 多次SRC" };

            left.Controls.Add(lblTitleMain);
            left.Controls.Add(lblMainLine1);
            left.Controls.Add(lblMainLine2);
            left.Controls.Add(lblTitleAux);
            left.Controls.Add(lblAuxLine1);
            left.Controls.Add(lblAuxLine2);

            // Right panel (settings)
            grpMain = new GroupBox { Left = 420, Top = 10, Width = 420, Height = 220, Text = "主通道设置" };
            grpAux  = new GroupBox { Left = 420, Top = 240, Width = 420, Height = 220, Text = "副通道设置" };
            this.Controls.Add(grpMain);
            this.Controls.Add(grpAux);

            // Main controls
            cboMainRate = MakeCombo(grpMain, "采样率", 10, 25, new[] {"192000","96000","48000","44100"});
            cboMainBits = MakeCombo(grpMain, "位深",   210, 25, new[] {"32","24","16"});
            nudMainBuf  = MakeNumeric(grpMain, "缓冲(ms)", 10, 70, 2, 1000, 12);
            cboMainShare= MakeCombo(grpMain, "共享/独占", 210, 70, new[] {"Auto","Exclusive","Shared"});
            cboMainSync = MakeCombo(grpMain, "同步模式", 10, 115, new[] {"Auto","Event","Polling"});
            cboMainAlign= MakeCombo(grpMain, "缓冲对齐", 210, 115, new[] {"DefaultAlign","MinAlign"});
            cboMainQuality = MakeCombo(grpMain, "程序内重采样质量", 10, 160, new[] {"60","50","40","30"});
            chkMainForceSharedResampler = MakeCheck(grpMain, "共享模式下也程序内重采样", 210, 160);

            // Aux controls
            cboAuxRate = MakeCombo(grpAux, "采样率", 10, 25, new[] {"48000","44100","96000","192000"});
            cboAuxBits = MakeCombo(grpAux, "位深",   210, 25, new[] {"16","24","32"});
            nudAuxBuf  = MakeNumeric(grpAux, "缓冲(ms)", 10, 70, 10, 1000, 150);
            cboAuxShare= MakeCombo(grpAux, "共享/独占", 210, 70, new[] {"Shared","Exclusive","Auto"});
            cboAuxSync = MakeCombo(grpAux, "同步模式", 10, 115, new[] {"Auto","Event","Polling"});
            cboAuxAlign= MakeCombo(grpAux, "缓冲对齐", 210, 115, new[] {"DefaultAlign","MinAlign"});
            cboAuxQuality = MakeCombo(grpAux, "程序内重采样质量", 10, 160, new[] {"30","40","50","60"});
            chkAuxForceSharedResampler = MakeCheck(grpAux, "共享模式下也程序内重采样", 210, 160);

            btnRefresh = new Button { Left = 420, Top = 470, Width = 90, Height = 28, Text = "刷新状态" };
            btnOk      = new Button { Left = 640, Top = 470, Width = 90, Height = 28, Text = "保存" };
            btnCancel  = new Button { Left = 740, Top = 470, Width = 90, Height = 28, Text = "取消" };
            this.Controls.AddRange(new Control[] { btnRefresh, btnOk, btnCancel });

            btnRefresh.Click += (s,e)=> RenderStatus();
            btnOk.Click += (s,e)=> { SaveToResult(); this.DialogResult = DialogResult.OK; this.Close(); };
            btnCancel.Click += (s,e)=> { this.DialogResult = DialogResult.Cancel; this.Close(); };
        }

        private ComboBox MakeCombo(Control parent, string title, int x, int y, string[] data)
        {
            var lab = new Label { Left = x, Top = y, Width = 180, Text = title };
            var cbo = new ComboBox { Left = x, Top = y + 18, Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            cbo.Items.AddRange(data);
            parent.Controls.Add(lab);
            parent.Controls.Add(cbo);
            return cbo;
        }

        private NumericUpDown MakeNumeric(Control parent, string title, int x, int y, int min, int max, int val)
        {
            var lab = new Label { Left = x, Top = y, Width = 180, Text = title };
            var nud = new NumericUpDown { Left = x, Top = y + 18, Width = 180, Minimum = min, Maximum = max, Value = val };
            parent.Controls.Add(lab);
            parent.Controls.Add(nud);
            return nud;
        }

        private CheckBox MakeCheck(Control parent, string title, int x, int y)
        {
            var chk = new CheckBox { Left = x, Top = y + 18, Width = 180, Text = title };
            parent.Controls.Add(chk);
            return chk;
        }

        private void LoadFromSettings()
        {
            // Main
            cboMainRate.SelectedItem = _settings.MainRate.ToString();
            cboMainBits.SelectedItem = _settings.MainBits.ToString();
            nudMainBuf.Value = _settings.MainBufMs;
            cboMainShare.SelectedItem = _settings.MainShare.ToString();
            cboMainSync.SelectedItem = _settings.MainSync.ToString();
            cboMainAlign.SelectedItem = _settings.MainBufMode.ToString();
            cboMainQuality.SelectedItem = _settings.MainResamplerQuality.ToString();
            chkMainForceSharedResampler.Checked = _settings.MainForceInternalResamplerInShared;

            // Aux
            cboAuxRate.SelectedItem = _settings.AuxRate.ToString();
            cboAuxBits.SelectedItem = _settings.AuxBits.ToString();
            nudAuxBuf.Value = _settings.AuxBufMs;
            cboAuxShare.SelectedItem = _settings.AuxShare.ToString();
            cboAuxSync.SelectedItem = _settings.AuxSync.ToString();
            cboAuxAlign.SelectedItem = _settings.AuxBufMode.ToString();
            cboAuxQuality.SelectedItem = _settings.AuxResamplerQuality.ToString();
            chkAuxForceSharedResampler.Checked = _settings.AuxForceInternalResamplerInShared;
        }

        private void SaveToResult()
        {
            Result = new AppSettings();
            // copy current _settings that we did not expose here if needed
            Result.InputDeviceId = _settings.InputDeviceId;
            Result.MainDeviceId = _settings.MainDeviceId;
            Result.AuxDeviceId = _settings.AuxDeviceId;
            Result.AutoStart = _settings.AutoStart;
            Result.EnableLogging = _settings.EnableLogging;
            Result.InputFormatStrategy = _settings.InputFormatStrategy;
            Result.InputCustomSampleRate = _settings.InputCustomSampleRate;
            Result.InputCustomBitDepth = _settings.InputCustomBitDepth;

            // Main
            int tmp;
            if (int.TryParse((string)cboMainRate.SelectedItem, out tmp)) Result.MainRate = tmp; else Result.MainRate = _settings.MainRate;
            if (int.TryParse((string)cboMainBits.SelectedItem, out tmp)) Result.MainBits = tmp; else Result.MainBits = _settings.MainBits;
            Result.MainBufMs = (int)nudMainBuf.Value;
            Result.MainShare = (ShareModeOption)Enum.Parse(typeof(ShareModeOption), (string)cboMainShare.SelectedItem);
            Result.MainSync  = (SyncModeOption)Enum.Parse(typeof(SyncModeOption), (string)cboMainSync.SelectedItem);
            Result.MainBufMode = (BufferAlignMode)Enum.Parse(typeof(BufferAlignMode), (string)cboMainAlign.SelectedItem);
            if (int.TryParse((string)cboMainQuality.SelectedItem, out tmp)) Result.MainResamplerQuality = tmp; else Result.MainResamplerQuality = _settings.MainResamplerQuality;
            Result.MainForceInternalResamplerInShared = chkMainForceSharedResampler.Checked;

            // Aux
            if (int.TryParse((string)cboAuxRate.SelectedItem, out tmp)) Result.AuxRate = tmp; else Result.AuxRate = _settings.AuxRate;
            if (int.TryParse((string)cboAuxBits.SelectedItem, out tmp)) Result.AuxBits = tmp; else Result.AuxBits = _settings.AuxBits;
            Result.AuxBufMs = (int)nudAuxBuf.Value;
            Result.AuxShare = (ShareModeOption)Enum.Parse(typeof(ShareModeOption), (string)cboAuxShare.SelectedItem);
            Result.AuxSync  = (SyncModeOption)Enum.Parse(typeof(SyncModeOption), (string)cboAuxSync.SelectedItem);
            Result.AuxBufMode = (BufferAlignMode)Enum.Parse(typeof(BufferAlignMode), (string)cboAuxAlign.SelectedItem);
            if (int.TryParse((string)cboAuxQuality.SelectedItem, out tmp)) Result.AuxResamplerQuality = tmp; else Result.AuxResamplerQuality = _settings.AuxResamplerQuality;
            Result.AuxForceInternalResamplerInShared = chkAuxForceSharedResampler.Checked;
        }

        private void RenderStatus()
        {
            var s = _statusProvider != null ? _statusProvider() : null;
            if (s == null)
            {
                lblMainLine1.Text = "主：状态不可用";
                lblMainLine2.Text = "主：程序内重采样质量=未生效 / 多次SRC=否";
                lblAuxLine1.Text  = "副：状态不可用";
                lblAuxLine2.Text  = "副：程序内重采样质量=未生效 / 多次SRC=否";
                // Disable right side if we don't have runtime
                SetMainEnable(false); SetAuxEnable(false);
                return;
            }

            // Line 1 stays compact: 直通/重采样、模式、格式、缓冲
            string mainResample = s.MainResampling ? "重采样=是" : "重采样=否";
            string auxResample  = s.AuxResampling ? "重采样=是" : "重采样=否";
            lblMainLine1.Text = string.Format("主：{0} | 模式={1} | 格式={2} | 缓冲={3}ms(对齐≈{4}×)",
                mainResample, s.MainMode, s.MainFormat, s.MainBufferMs, s.MainAlignedMultiple);
            lblAuxLine1.Text = string.Format("副：{0} | 模式={1} | 格式={2} | 缓冲={3}ms(对齐≈{4}×)",
                auxResample, s.AuxMode, s.AuxFormat, s.AuxBufferMs, s.AuxAlignedMultiple);

            // Line 2: 程序内重采样质量 + 多次SRC
            string mainQ = s.MainInternalResampler ? s.MainInternalResamplerQuality.ToString() : "未生效";
            string auxQ  = s.AuxInternalResampler  ? s.AuxInternalResamplerQuality.ToString()  : "未生效";
            lblMainLine2.Text = string.Format("主：程序内重采样={0} / 质量={1} / 多次SRC={2}",
                s.MainInternalResampler ? "是" : "否", mainQ, s.MainMultiSRC ? "是" : "否");
            lblAuxLine2.Text = string.Format("副：程序内重采样={0} / 质量={1} / 多次SRC={2}",
                s.AuxInternalResampler ? "是" : "否", auxQ, s.AuxMultiSRC ? "是" : "否");

            // Disable logic：当“程序内重采样=否”时 -> 质量下拉与“共享模式下也程序内重采样”灰掉
            SetMainEnable(s.MainInternalResampler);
            SetAuxEnable(s.AuxInternalResampler);
        }

        private void SetMainEnable(bool internalActive)
        {
            cboMainQuality.Enabled = internalActive;
            chkMainForceSharedResampler.Enabled = internalActive;
        }
        private void SetAuxEnable(bool internalActive)
        {
            cboAuxQuality.Enabled = internalActive;
            chkAuxForceSharedResampler.Enabled = internalActive;
        }
    }
}
