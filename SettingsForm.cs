using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MirrorAudio
{
    public partial class SettingsForm : Form
    {
        private readonly Func<StatusSnapshot> _getStatus;
        private readonly AppSettings _cfg;

        // —— UI 控件（仅列与本补丁相关的关键控件）——
        ComboBox cmbInput, cmbMain, cmbAux;
        ComboBox cmbShareMain, cmbShareAux, cmbSyncMain, cmbSyncAux, cmbBufModeMain, cmbBufModeAux;
        NumericUpDown numBufMain, numBufAux, numRateMain, numBitsMain, numRateAux, numBitsAux;
        ComboBox cmbResampMain, cmbResampAux;
        CheckBox chkMainForceInShared, chkAuxForceInShared;
        CheckBox chkAutoStart, chkLogging;
        ComboBox cmbInStrategy; NumericUpDown numInRate, numInBits;

        Label lblMainResampHint, lblAuxResampHint;
        Button btnOk, btnCancel, btnApply, btnInit;

        public AppSettings Result { get; private set; }

        // 对外事件：“应用(不关闭)”
        public event Action<AppSettings> ApplyRequested;

        public SettingsForm(AppSettings cfg, Func<StatusSnapshot> getStatus)
        {
            _cfg = cfg; // 由上层决定是否 Clone
            _getStatus = getStatus;

            InitializeComponent();
            BuildUi();
            LoadFromConfig(_cfg);
            RefreshResamplerQualityEnabled();
        }

        private void InitializeComponent()
        {
            this.Text = "MirrorAudio 设置";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(880, 640);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            this.Controls.Add(root);

            // —— 顶部：输入配置（录音/环回只共享）——
            var grpInput = new GroupBox { Text = "输入（固定共享）", Dock = DockStyle.Top, Padding = new Padding(10) };
            var pIn = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 2 };
            grpInput.Controls.Add(pIn);
            root.Controls.Add(grpInput);

            cmbInput = NewCombo();
            cmbInStrategy = NewCombo();
            cmbInStrategy.Items.AddRange(new object[] { "系统混音", "自定义", "32f优先" });
            numInRate = NewNum(8000, 768000, 48000);
            numInBits = NewNum(8, 32, 24);

            AddRow(pIn,
                new Label { Text = "输入设备：" }, cmbInput,
                new Label { Text = "输入策略：" }, cmbInStrategy,
                new Label { Text = "自定义SR/Bit：" }, HStack(numInRate, new Label { Text = "Hz   " }, numInBits, new Label { Text = "bit" })
            );

            // —— 中部：主/副输出完全对称 —— 
            var grpOut = new GroupBox { Text = "输出", Dock = DockStyle.Fill, Padding = new Padding(10) };
            root.Controls.Add(grpOut);
            var pOut = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 8, AutoScroll = true };
            grpOut.Controls.Add(pOut);

            // 主
            cmbMain = NewCombo();
            cmbShareMain = NewCombo(); cmbShareMain.Items.AddRange(new object[] { "自动", "独占", "共享" });
            cmbSyncMain = NewCombo(); cmbSyncMain.Items.AddRange(new object[] { "自动", "事件", "轮询" });
            cmbBufModeMain = NewCombo(); cmbBufModeMain.Items.AddRange(new object[] { "默认对齐", "最小对齐" });
            numBufMain = NewNum(3, 500, 12);
            numRateMain = NewNum(8000, 768000, 48000);
            numBitsMain = NewNum(8, 32, 24);
            cmbResampMain = NewCombo(); cmbResampMain.Items.AddRange(new object[] { "60", "50", "40", "30" });
            chkMainForceInShared = new CheckBox { Text = "共享模式下也使用内部重采样" };
            lblMainResampHint = new Label { Text = "当前由系统混音/SRC，该设置不生效", ForeColor = Color.DimGray, AutoSize = true, Visible = false };

            // 副
            cmbAux = NewCombo();
            cmbShareAux = NewCombo(); cmbShareAux.Items.AddRange(new object[] { "自动", "独占", "共享" });
            cmbSyncAux = NewCombo(); cmbSyncAux.Items.AddRange(new object[] { "自动", "事件", "轮询" });
            cmbBufModeAux = NewCombo(); cmbBufModeAux.Items.AddRange(new object[] { "默认对齐", "最小对齐" });
            numBufAux = NewNum(3, 800, 150);
            numRateAux = NewNum(8000, 768000, 48000);
            numBitsAux = NewNum(8, 32, 24);
            cmbResampAux = NewCombo(); cmbResampAux.Items.AddRange(new object[] { "60", "50", "40", "30" });
            chkAuxForceInShared = new CheckBox { Text = "共享模式下也使用内部重采样" };
            lblAuxResampHint = new Label { Text = "当前由系统混音/SRC，该设置不生效", ForeColor = Color.DimGray, AutoSize = true, Visible = false };

            // 主路行
            AddRow(pOut, new Label { Text = "主输出设备：" }, cmbMain, new Label { Text = "模式：" }, cmbShareMain, new Label { Text = "同步：" }, cmbSyncMain);
            AddRow(pOut, new Label { Text = "主缓冲(ms)：" }, numBufMain, new Label { Text = "缓冲对齐：" }, cmbBufModeMain, new Label { Text = "SR/Bit：" }, HStack(numRateMain, new Label { Text = "Hz   " }, numBitsMain, new Label { Text = "bit" }));
            AddRow(pOut, new Label { Text = "重采样质量：" }, HStack(cmbResampMain, lblMainResampHint), new Label { Text = "" }, chkMainForceInShared, new Label { Text = "" }, new Label { Text = "" });

            // 副路行
            AddRow(pOut, new Label { Text = "副输出设备：" }, cmbAux, new Label { Text = "模式：" }, cmbShareAux, new Label { Text = "同步：" }, cmbSyncAux);
            AddRow(pOut, new Label { Text = "副缓冲(ms)：" }, numBufAux, new Label { Text = "缓冲对齐：" }, cmbBufModeAux, new Label { Text = "SR/Bit：" }, HStack(numRateAux, new Label { Text = "Hz   " }, numBitsAux, new Label { Text = "bit" }));
            AddRow(pOut, new Label { Text = "重采样质量：" }, HStack(cmbResampAux, lblAuxResampHint), new Label { Text = "" }, chkAuxForceInShared, new Label { Text = "" }, new Label { Text = "" });

            // —— 底部：其他开关 —— 
            var grpOther = new GroupBox { Text = "其他", Dock = DockStyle.Top, Padding = new Padding(10) };
            root.Controls.Add(grpOther);
            var pO = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            grpOther.Controls.Add(pO);
            chkAutoStart = new CheckBox { Text = "开机自启动" };
            chkLogging = new CheckBox { Text = "启用日志" };
            pO.Controls.Add(chkAutoStart);
            pO.Controls.Add(chkLogging);

            // —— 按钮区 —— 
            var pBtn = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10) };
            root.Controls.Add(pBtn);
            btnOk = new Button { Text = "保存并关闭", AutoSize = true };
            btnCancel = new Button { Text = "取消", AutoSize = true };
            btnApply = new Button { Text = "应用(不关闭)", AutoSize = true };
            btnInit = new Button { Text = "程序初始化", AutoSize = true };

            pBtn.Controls.AddRange(new Control[] { btnOk, btnCancel, btnApply, btnInit });

            btnOk.Click += (_, __) => { Result = BuildCurrentSettings(); DialogResult = DialogResult.OK; Close(); };
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            btnApply.Click += (_, __) => { ApplyRequested?.Invoke(BuildCurrentSettings()); RefreshResamplerQualityEnabled(); };
            btnInit.Click += (_, __) => { Preset_Main_LowLatencyLossless(); Preset_Aux_Standard(); };

            // 占位：你项目应替换为实际设备枚举
            cmbInput.Items.Add("System Default Input");
            cmbMain .Items.Add("System Default Output (Main)");
            cmbAux  .Items.Add("System Default Output (Aux)");
            cmbInput.SelectedIndex = 0; cmbMain.SelectedIndex = 0; cmbAux.SelectedIndex = 0;
        }

        private static ComboBox NewCombo() => new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        private static NumericUpDown NewNum(int min, int max, int val) => new NumericUpDown { Minimum = min, Maximum = max, Value = val, Increment = 1, Dock = DockStyle.Fill };
        private static Control HStack(params Control[] cs)
        {
            var p = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
            foreach (var c in cs) p.Controls.Add(c);
            return p;
        }
        private static void AddRow(TableLayoutPanel t, params Control[] cs)
        {
            var r = t.RowCount++;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.ColumnStyles.Clear();
            for (int i = 0; i < t.ColumnCount; i++) t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / t.ColumnCount));
            for (int i = 0; i < cs.Length; i++) t.Controls.Add(cs[i], i, r);
        }

        private void LoadFromConfig(AppSettings c)
        {
            // 输入
            cmbInStrategy.SelectedIndex = (int)c.InputFormatStrategy;
            numInRate.Value = c.InputCustomSampleRate;
            numInBits.Value = c.InputCustomBitDepth;

            // 主
            cmbShareMain.SelectedIndex = ToIdx(c.MainShare);
            cmbSyncMain .SelectedIndex = ToIdx(c.MainSync);
            cmbBufModeMain.SelectedIndex = (c.MainBufMode == BufferAlignMode.MinAlign) ? 1 : 0;
            numBufMain.Value = c.MainBufMs;
            numRateMain.Value = c.MainRate;
            numBitsMain.Value = c.MainBits;
            cmbResampMain.SelectedItem = c.MainResamplerQuality.ToString();
            chkMainForceInShared.Checked = c.MainForceInternalResamplerInShared;

            // 副
            cmbShareAux.SelectedIndex = ToIdx(c.AuxShare);
            cmbSyncAux .SelectedIndex = ToIdx(c.AuxSync);
            cmbBufModeAux.SelectedIndex = (c.AuxBufMode == BufferAlignMode.MinAlign) ? 1 : 0;
            numBufAux.Value = c.AuxBufMs;
            numRateAux.Value = c.AuxRate;
            numBitsAux.Value = c.AuxBits;
            cmbResampAux.SelectedItem = c.AuxResamplerQuality.ToString();
            chkAuxForceInShared.Checked = c.AuxForceInternalResamplerInShared;

            // 其他
            chkAutoStart.Checked = c.AutoStart;
            chkLogging.Checked = c.EnableLogging;
        }

        private int ToIdx(ShareModeOption s) => s == ShareModeOption.Exclusive ? 1 : s == ShareModeOption.Shared ? 2 : 0;
        private int ToIdx(SyncModeOption s) => s == SyncModeOption.Event ? 1 : s == SyncModeOption.Polling ? 2 : 0;

        private AppSettings BuildCurrentSettings()
        {
            var s = _cfg; // 直接回写

            // 输入
            s.InputFormatStrategy = (InputFormatStrategy)cmbInStrategy.SelectedIndex;
            s.InputCustomSampleRate = (int)numInRate.Value;
            s.InputCustomBitDepth = (int)numInBits.Value;

            // 主
            s.MainShare = cmbShareMain.SelectedIndex == 1 ? ShareModeOption.Exclusive : (cmbShareMain.SelectedIndex == 2 ? ShareModeOption.Shared : ShareModeOption.Auto);
            s.MainSync  = cmbSyncMain .SelectedIndex == 1 ? SyncModeOption.Event    : (cmbSyncMain .SelectedIndex == 2 ? SyncModeOption.Polling  : SyncModeOption.Auto);
            s.MainBufMode = (cmbBufModeMain.SelectedIndex == 1) ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign;
            s.MainBufMs = (int)numBufMain.Value;
            s.MainRate = (int)numRateMain.Value; // 不动 SR/位深时，初始化逻辑会覆盖延迟策略但不改这两个
            s.MainBits = (int)numBitsMain.Value;
            s.MainResamplerQuality = int.Parse((string)cmbResampMain.SelectedItem);
            s.MainForceInternalResamplerInShared = chkMainForceInShared.Checked;

            // 副
            s.AuxShare = cmbShareAux.SelectedIndex == 1 ? ShareModeOption.Exclusive : (cmbShareAux.SelectedIndex == 2 ? ShareModeOption.Shared : ShareModeOption.Auto);
            s.AuxSync  = cmbSyncAux .SelectedIndex == 1 ? SyncModeOption.Event     : (cmbSyncAux .SelectedIndex == 2 ? SyncModeOption.Polling  : SyncModeOption.Auto);
            s.AuxBufMode = (cmbBufModeAux.SelectedIndex == 1) ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign;
            s.AuxBufMs = (int)numBufAux.Value;
            s.AuxRate = (int)numRateAux.Value;
            s.AuxBits = (int)numBitsAux.Value;
            s.AuxResamplerQuality = int.Parse((string)cmbResampAux.SelectedItem);
            s.AuxForceInternalResamplerInShared = chkAuxForceInShared.Checked;

            // 其他
            s.AutoStart = chkAutoStart.Checked;
            s.EnableLogging = chkLogging.Checked;

            Result = s;
            return s;
        }

        private void Preset_Main_LowLatencyLossless()
        {
            cmbShareMain.SelectedIndex = 1; // 独占
            cmbSyncMain .SelectedIndex = 1; // 事件
            cmbBufModeMain.SelectedIndex = 1; // 最小对齐
            numBufMain.Value = Math.Min(Math.Max(12, (int)numBufMain.Minimum), (int)numBufMain.Maximum);
            cmbResampMain.SelectedItem = "50";
            chkMainForceInShared.Checked = false; // 共享下不强制程序内重采样
        }

        private void Preset_Aux_Standard()
        {
            cmbShareAux.SelectedIndex = 2; // 共享
            cmbSyncAux .SelectedIndex = 0; // 自动
            cmbBufModeAux.SelectedIndex = 0; // 默认对齐
            numBufAux.Value = Math.Min(Math.Max(150, (int)numBufAux.Minimum), (int)numBufAux.Maximum);
            cmbResampAux.SelectedItem = "30";
            chkAuxForceInShared.Checked = false;
        }

        private void RefreshResamplerQualityEnabled()
        {
            var st = _getStatus?.Invoke();

            bool mainEnable = st?.MainInternalResampler ?? false;
            bool auxEnable  = st?.AuxInternalResampler  ?? false;

            cmbResampMain.Enabled = mainEnable;
            lblMainResampHint.Visible = !mainEnable;

            cmbResampAux.Enabled = auxEnable;
            lblAuxResampHint.Visible = !auxEnable;
        }
    }
}

        private void RefreshStatusLabels()
        {
            var st = _getStatus?.Invoke();
            if (st != null)
            {
                lblStatusMainInternal.Text = (st.MainInternalResampler ? "是" : "否");
                lblStatusMainMulti.Text    = (st.MainMultiSRC ? "是" : "否");
                lblStatusAuxInternal.Text  = (st.AuxInternalResampler ? "是" : "否");
                lblStatusAuxMulti.Text     = (st.AuxMultiSRC ? "是" : "否");
            }
        }
