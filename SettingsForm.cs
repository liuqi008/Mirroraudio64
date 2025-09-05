using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    sealed class SettingsForm : Form
    {
        // 左：状态
        readonly Label lblRun = new Label();
        readonly Label lblInput = new Label();
        readonly Label lblMain  = new Label();
        readonly Label lblAux   = new Label();
        readonly Label lblMainFmt = new Label();
        readonly Label lblAuxFmt  = new Label();
        readonly Label lblMainBuf = new Label();
        readonly Label lblAuxBuf  = new Label();
        readonly Label lblMainPer = new Label();
        readonly Label lblAuxPer  = new Label();

        // 右：设置
        readonly ComboBox cmbInput = new ComboBox();
        readonly ComboBox cmbMain  = new ComboBox();
        readonly ComboBox cmbAux   = new ComboBox();

        readonly ComboBox cmbShareMain = new ComboBox(); // 主：共享/独占/自动
        readonly ComboBox cmbSyncMain  = new ComboBox(); // 主：事件/轮询/自动

        readonly ComboBox cmbShareAux  = new ComboBox(); // 副：共享/独占/自动
        readonly ComboBox cmbSyncAux   = new ComboBox(); // 副：事件/轮询/自动

        readonly NumericUpDown numRateMain = new NumericUpDown();
        readonly NumericUpDown numBitsMain = new NumericUpDown();
        readonly NumericUpDown numBufMain  = new NumericUpDown();

        readonly NumericUpDown numRateAux  = new NumericUpDown();
        readonly NumericUpDown numBitsAux  = new NumericUpDown();
        readonly NumericUpDown numBufAux   = new NumericUpDown();

        readonly CheckBox chkAutoStart = new CheckBox();
        readonly CheckBox chkLogging   = new CheckBox();

        readonly Button btnOk = new Button();
        readonly Button btnCancel = new Button();

        readonly Button btnRefresh = new Button();
        readonly Button btnCopy    = new Button();
        readonly Button btnReload  = new Button();

        readonly Func<StatusSnapshot> _statusProvider;

        sealed class DevItem { public string Id; public string Name; public override string ToString() { return Name; } }

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings cur, Func<StatusSnapshot> statusProvider)
        {
            _statusProvider = statusProvider ?? (() => new StatusSnapshot{ Running=false });

            Text = "MirrorAudio 设置";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = SystemFonts.MessageBoxFont;
            MinimumSize = new Size(980, 640);
            Size = new Size(1100, 720);

            // 顶层分栏：左状态、右设置
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, FixedPanel = FixedPanel.Panel1, SplitterWidth = 6 };
            split.Panel1MinSize = 360;
            split.SplitterDistance = 380;

            // —— 左侧：状态卡片 —— //
            var leftScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(12) };
            var grpStatus = new GroupBox { Text = "当前状态（打开查看，关闭即释放内存）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) };
            var tblS = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
            tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));

            AddRow(tblS, "运行状态", lblRun);
            AddRow(tblS, "输入",    lblInput);
            AddRow(tblS, "主通道",  lblMain);
            AddRow(tblS, "主格式",  lblMainFmt);
            AddRow(tblS, "主缓冲",  lblMainBuf);
            AddRow(tblS, "主周期",  lblMainPer);
            AddRow(tblS, "副通道",  lblAux);
            AddRow(tblS, "副格式",  lblAuxFmt);
            AddRow(tblS, "副缓冲",  lblAuxBuf);
            AddRow(tblS, "副周期",  lblAuxPer);

            var pnlSBtns = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 4) };
            btnRefresh.Text = "刷新状态";
            btnCopy.Text = "复制状态";
            btnRefresh.Click += (s,e)=> RenderStatus();
            btnCopy.Click += (s,e)=> { Clipboard.SetText(BuildStatusText()); MessageBox.Show("状态已复制到剪贴板。","MirrorAudio",MessageBoxButtons.OK,MessageBoxIcon.Information); };
            pnlSBtns.Controls.Add(btnRefresh);
            pnlSBtns.Controls.Add(btnCopy);
            grpStatus.Controls.Add(tblS);
            grpStatus.Controls.Add(pnlSBtns);

            leftScroll.Controls.Add(grpStatus);
            split.Panel1.Controls.Add(leftScroll);

            // —— 右侧：设置 —— //
            var rightScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(12, 8, 12, 8) };

            var grpDevices = new GroupBox { Text = "设备选择", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) };
            var gridDev = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            gridDev.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            gridDev.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            foreach (var cb in new[] { cmbInput, cmbMain, cmbAux }) cb.DropDownStyle = ComboBoxStyle.DropDownList;
            AddRow(gridDev, "通道1 输入（录音/环回）", cmbInput);
            AddRow(gridDev, "通道2 主通道（低延迟）",   cmbMain);
            AddRow(gridDev, "通道3 副通道（推流）",     cmbAux);
            btnReload.Text = "重新枚举设备";
            btnReload.AutoSize = true;
            btnReload.Click += (s,e)=> LoadDevices();
            gridDev.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            gridDev.Controls.Add(btnReload, 1, gridDev.RowCount++);
            grpDevices.Controls.Add(gridDev);

            var grpMain = new GroupBox { Text = "主通道（低延迟）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) };
            var gridMain = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            gridMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            gridMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            foreach (var cb in new[] { cmbShareMain, cmbSyncMain }) cb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbShareMain.Items.AddRange(new object[] { "自动（优先独占）", "强制独占", "强制共享" });
            cmbSyncMain.Items.AddRange(new object[]  { "自动（事件优先）", "强制事件", "强制轮询" });
            AddRow(gridMain, "模式", cmbShareMain);
            AddRow(gridMain, "同步方式", cmbSyncMain);

            numRateMain.Maximum = 384000; numRateMain.Minimum = 44100; numRateMain.Increment = 1000;
            numBitsMain.Maximum = 32;     numBitsMain.Minimum = 16;    numBitsMain.Increment = 8;
            numBufMain.Maximum  = 200;    numBufMain.Minimum  = 4;
            AddRow(gridMain, "采样率 (Hz，仅独占生效)",   numRateMain);
            AddRow(gridMain, "位深 (bit，仅独占生效)",     numBitsMain);
            AddRow(gridMain, "缓冲 (ms)",                   numBufMain);
            grpMain.Controls.Add(gridMain);

            var grpAux = new GroupBox { Text = "副通道（推流）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) };
            var gridAux = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            gridAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            gridAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            foreach (var cb in new[] { cmbShareAux, cmbSyncAux }) cb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbShareAux.Items.AddRange(new object[] { "自动（优先独占）", "强制独占", "强制共享" });
            cmbSyncAux.Items.AddRange(new object[]  { "自动（事件优先）", "强制事件", "强制轮询" });
            AddRow(gridAux, "模式", cmbShareAux);
            AddRow(gridAux, "同步方式", cmbSyncAux);

            numRateAux.Maximum = 384000;  numRateAux.Minimum = 44100; numRateAux.Increment = 1000;
            numBitsAux.Maximum = 32;      numBitsAux.Minimum = 16;    numBitsAux.Increment = 8;
            numBufAux.Maximum  = 400;     numBufAux.Minimum  = 50;
            AddRow(gridAux, "采样率 (Hz，仅独占生效)",  numRateAux);
            AddRow(gridAux, "位深 (bit，仅独占生效)",   numBitsAux);
            AddRow(gridAux, "缓冲 (ms)",                 numBufAux);
            grpAux.Controls.Add(gridAux);

            var grpOpts = new GroupBox { Text = "其他选项", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) };
            var pnlOpts = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, AutoSize = true };
            chkAutoStart.Text = "Windows 自启动";
            chkLogging.Text   = "启用日志（排障用）";
            pnlOpts.Controls.Add(chkAutoStart);
            pnlOpts.Controls.Add(chkLogging);
            grpOpts.Controls.Add(pnlOpts);

            var tip = new Label {
                Text = "提示：共享模式由系统混音决定格式；24-bit 独占失败时，可尝试 32-bit（24-in-32 容器）。",
                AutoSize = true, ForeColor = SystemColors.GrayText, Padding = new Padding(4, 4, 0, 4)
            };

            // 右侧布局序列
            rightScroll.Controls.Add(grpOpts);
            rightScroll.Controls.Add(tip);
            rightScroll.Controls.Add(grpAux);
            rightScroll.Controls.Add(grpMain);
            rightScroll.Controls.Add(grpDevices);

            split.Panel2.Controls.Add(rightScroll);

            // 底部按钮栏（右对齐）
            var pnlButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Padding = new Padding(12), AutoSize = true };
            btnOk.Text = "保存"; btnCancel.Text = "取消";
            AcceptButton = btnOk; CancelButton = btnCancel;
            btnOk.DialogResult = DialogResult.OK; btnCancel.DialogResult = DialogResult.Cancel;
            btnOk.Click += (s,e)=> SaveAndClose();
            pnlButtons.Controls.AddRange(new Control[]{ btnOk, btnCancel });

            Controls.Add(split);
            Controls.Add(pnlButtons);

            // 载入设备与配置
            LoadDevices();
            LoadConfig(cur);

            // 初次渲染状态
            RenderStatus();
        }

        void LoadConfig(AppSettings cur)
        {
            Result = new AppSettings {
                InputDeviceId = cur.InputDeviceId, MainDeviceId = cur.MainDeviceId, AuxDeviceId = cur.AuxDeviceId,
                MainShare = cur.MainShare, MainSync = cur.MainSync,
                MainRate  = cur.MainRate,  MainBits  = cur.MainBits,  MainBufMs = cur.MainBufMs,
                AuxShare  = cur.AuxShare,  AuxSync   = cur.AuxSync,
                AuxRate   = cur.AuxRate,   AuxBits   = cur.AuxBits,   AuxBufMs  = cur.AuxBufMs,
                AutoStart = cur.AutoStart, EnableLogging = cur.EnableLogging
            };

            SelectById(cmbInput, cur.InputDeviceId);
            SelectById(cmbMain,  cur.MainDeviceId);
            SelectById(cmbAux,   cur.AuxDeviceId);

            numRateMain.Value = Clamp(cur.MainRate, (int)numRateMain.Minimum, (int)numRateMain.Maximum);
            numBitsMain.Value = Clamp(cur.MainBits, (int)numBitsMain.Minimum, (int)numBitsMain.Maximum);
            numBufMain.Value  = Clamp(cur.MainBufMs,(int)numBufMain.Minimum,(int)numBufMain.Maximum);

            cmbShareMain.SelectedIndex = cur.MainShare == ShareModeOption.Auto ? 0 : (cur.MainShare == ShareModeOption.Exclusive ? 1 : 2);
            cmbSyncMain.SelectedIndex  = cur.MainSync  == SyncModeOption.Auto  ? 0 : (cur.MainSync  == SyncModeOption.Event     ? 1 : 2);

            numRateAux.Value = Clamp(cur.AuxRate, (int)numRateAux.Minimum, (int)numRateAux.Maximum);
            numBitsAux.Value = Clamp(cur.AuxBits, (int)numBitsAux.Minimum, (int)numBitsAux.Maximum);
            numBufAux.Value  = Clamp(cur.AuxBufMs,(int)numBufAux.Minimum,(int)numBufAux.Maximum);

            cmbShareAux.SelectedIndex = cur.AuxShare == ShareModeOption.Auto ? 0 : (cur.AuxShare == ShareModeOption.Exclusive ? 1 : 2);
            cmbSyncAux.SelectedIndex  = cur.AuxSync  == SyncModeOption.Auto  ? 0 : (cur.AuxSync  == SyncModeOption.Event     ? 1 : 2);

            chkAutoStart.Checked = cur.AutoStart;
            chkLogging.Checked   = cur.EnableLogging;
        }

        void RenderStatus()
        {
            StatusSnapshot s;
            try { s = _statusProvider(); } catch { s = new StatusSnapshot(); }

            lblRun.Text    = s.Running ? "运行中" : "停止";
            lblInput.Text  = (s.InputDevice ?? "-") + " | " + (s.InputRole ?? "-") + " | " + (s.InputFormat ?? "-");
            lblMain.Text   = (s.MainDevice  ?? "-") + " | " + (s.MainMode  ?? "-") + " | " + (s.MainSync ?? "-");
            lblAux.Text    = (s.AuxDevice   ?? "-") + " | " + (s.AuxMode   ?? "-") + " | " + (s.AuxSync  ?? "-");
            lblMainFmt.Text= s.MainFormat ?? "-";
            lblAuxFmt.Text = s.AuxFormat  ?? "-";
            lblMainBuf.Text= s.MainBufferMs > 0 ? (s.MainBufferMs + " ms") : "-";
            lblAuxBuf.Text = s.AuxBufferMs  > 0 ? (s.AuxBufferMs  + " ms") : "-";
            lblMainPer.Text= "默认 " + s.MainDefaultPeriodMs.ToString("0.##") + " ms / 最小 " + s.MainMinimumPeriodMs.ToString("0.##") + " ms";
            lblAuxPer.Text = "默认 " + s.AuxDefaultPeriodMs.ToString("0.##")  + " ms / 最小 " + s.AuxMinimumPeriodMs.ToString("0.##")  + " ms";
        }

        string BuildStatusText()
        {
            StatusSnapshot s;
            try { s = _statusProvider(); } catch { s = new StatusSnapshot(); }

            var sb = new StringBuilder(256);
            sb.AppendLine("MirrorAudio 状态快照");
            sb.AppendLine("运行状态: " + (s.Running ? "运行中" : "停止"));
            sb.AppendLine("输入: " + (s.InputDevice ?? "-") + " | " + (s.InputRole ?? "-") + " | " + (s.InputFormat ?? "-"));
            sb.AppendLine("主通道: " + (s.MainDevice ?? "-") + " | " + (s.MainMode ?? "-") + " | " + (s.MainSync ?? "-"));
            sb.AppendLine("主格式: " + (s.MainFormat ?? "-"));
            sb.AppendLine("主缓冲: " + (s.MainBufferMs > 0 ? (s.MainBufferMs + " ms") : "-"));
            sb.AppendLine("主周期: 默认 " + s.MainDefaultPeriodMs.ToString("0.##") + " ms / 最小 " + s.MainMinimumPeriodMs.ToString("0.##") + " ms");
            sb.AppendLine("副通道: " + (s.AuxDevice ?? "-") + " | " + (s.AuxMode ?? "-") + " | " + (s.AuxSync ?? "-"));
            sb.AppendLine("副格式: " + (s.AuxFormat ?? "-"));
            sb.AppendLine("副缓冲: " + (s.AuxBufferMs > 0 ? (s.AuxBufferMs + " ms") : "-"));
            sb.AppendLine("副周期: 默认 " + s.AuxDefaultPeriodMs.ToString("0.##") + " ms / 最小 " + s.AuxMinimumPeriodMs.ToString("0.##") + " ms");
            return sb.ToString();
        }

        static void AddRow(TableLayoutPanel grid, string label, Control c)
        {
            var l = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Padding = new Padding(0,6,0,6) };
            c.Dock = DockStyle.Fill;
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grid.Controls.Add(l, 0, grid.RowCount);
            grid.Controls.Add(c, 1, grid.RowCount);
            grid.RowCount++;
        }

        static int Clamp(int v, int lo, int hi) { return v < lo ? lo : (v > hi ? hi : v); }

        void LoadDevices()
        {
            var mm = new MMDeviceEnumerator();
            cmbInput.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="录音: " + d.FriendlyName});
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="环回: " + d.FriendlyName});

            cmbMain.Items.Clear(); cmbAux.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                var it = new DevItem{ Id=d.ID, Name=d.FriendlyName };
                cmbMain.Items.Add(it);
                cmbAux.Items.Add(new DevItem{ Id=d.ID, Name=d.FriendlyName });
            }
        }

        void SelectById(ComboBox cmb, string id)
        {
            if (string.IsNullOrEmpty(id) || cmb.Items.Count == 0) { if (cmb.Items.Count>0) cmb.SelectedIndex=0; return; }
            for (int i=0;i<cmb.Items.Count;i++)
            {
                var it = cmb.Items[i] as DevItem;
                if (it != null && it.Id == id) { cmb.SelectedIndex = i; return; }
            }
            cmb.SelectedIndex = 0;
        }

        void SaveAndClose()
        {
            var inSel  = cmbInput.SelectedItem as DevItem;
            var mainSel= cmbMain.SelectedItem  as DevItem;
            var auxSel = cmbAux.SelectedItem   as DevItem;
            if (mainSel == null || auxSel == null)
            {
                MessageBox.Show("请至少选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None; return;
            }

            ShareModeOption shareMain = cmbShareMain.SelectedIndex == 1 ? ShareModeOption.Exclusive : (cmbShareMain.SelectedIndex == 2 ? ShareModeOption.Shared : ShareModeOption.Auto);
            SyncModeOption  syncMain  = cmbSyncMain.SelectedIndex  == 1 ? SyncModeOption.Event      : (cmbSyncMain.SelectedIndex  == 2 ? SyncModeOption.Polling : SyncModeOption.Auto);

            ShareModeOption shareAux  = cmbShareAux.SelectedIndex == 1 ? ShareModeOption.Exclusive : (cmbShareAux.SelectedIndex == 2 ? ShareModeOption.Shared : ShareModeOption.Auto);
            SyncModeOption  syncAux   = cmbSyncAux.SelectedIndex  == 1 ? SyncModeOption.Event      : (cmbSyncAux.SelectedIndex  == 2 ? SyncModeOption.Polling : SyncModeOption.Auto);

            Result = new AppSettings {
                InputDeviceId = inSel != null ? inSel.Id : null,
                MainDeviceId  = mainSel.Id,
                AuxDeviceId   = auxSel.Id,

                MainShare     = shareMain,
                MainSync      = syncMain,
                MainRate      = (int)numRateMain.Value,
                MainBits      = (int)numBitsMain.Value,
                MainBufMs     = (int)numBufMain.Value,

                AuxShare      = shareAux,
                AuxSync       = syncAux,
                AuxRate       = (int)numRateAux.Value,
                AuxBits       = (int)numBitsAux.Value,
                AuxBufMs      = (int)numBufAux.Value,

                AutoStart     = chkAutoStart.Checked,
                EnableLogging = chkLogging.Checked
            };
            DialogResult = DialogResult.OK;
            Close(); // 关闭即释放窗口内存
        }
    }
}
