using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    public sealed class SettingsForm : Form
    {
        // ====== 对外 ======
        public AppSettings Result { get; private set; }
        private readonly AppSettings _orig;
        private readonly Func<StatusSnapshot> _getStatus;

        // ====== 设备枚举 ======
        private readonly MMDeviceEnumerator _mm = new MMDeviceEnumerator();

        // ====== 左：状态控件 ======
        Label lblRunning, lblInDev, lblInRole, lblInFmt;
        Label lblMainDev, lblMainMode, lblMainSync, lblMainFmt, lblMainBuf, lblMainPass, lblMainPeriod;
        Label lblAuxDev, lblAuxMode, lblAuxSync, lblAuxFmt, lblAuxBuf, lblAuxPass, lblAuxPeriod, lblAuxQ;
        Button btnRefreshStatus;

        // ====== 右：设备与设置控件 ======
        ComboBox cmbIn, cmbMain, cmbAux;
        Button btnRescan;

        ComboBox cmbMainShare, cmbMainSync, cmbMainBits, cmbMainRate;
        NumericUpDown nudMainBuf;

        ComboBox cmbAuxShare, cmbAuxSync, cmbAuxBits, cmbAuxRate, cmbAuxQuality;
        NumericUpDown nudAuxBuf;

        CheckBox chkAutoStart, chkLogging, chkForceUserFormat, chkForcePassthroughStrict;

        Button btnOK, btnCancel;

        public SettingsForm(AppSettings cfg, Func<StatusSnapshot> getStatus)
        {
            _orig = cfg ?? new AppSettings();
            _getStatus = getStatus;

            Text = "MirrorAudio 设置";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(980, 560);
            MaximizeBox = false;

            // ===== 根：左右两栏 =====
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.Panel1,
                SplitterWidth = 6,
                IsSplitterFixed = false
            };
            split.Panel1MinSize = 420;   // 左：状态
            split.Panel2MinSize = 520;   // 右：设置
            Controls.Add(split);

            // ===== 左：状态 - 用表格严格布局 =====
            var left = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(12),
                AutoScroll = true
            };
            left.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            split.Panel1.Controls.Add(left);

            int r = 0;
            AddHeader(left, ref r, "当前状态");
            lblRunning = AddKV(left, ref r, "运行中：");
            lblInDev = AddKV(left, ref r, "输入设备：");
            lblInRole = AddKV(left, ref r, "输入角色：");
            lblInFmt = AddKV(left, ref r, "输入格式：");

            AddSeparator(left, ref r);
            AddHeader(left, ref r, "主输出（高音质 / 低延迟）");
            lblMainDev = AddKV(left, ref r, "设备：");
            lblMainMode = AddKV(left, ref r, "模式：");
            lblMainSync = AddKV(left, ref r, "同步：");
            lblMainFmt = AddKV(left, ref r, "实际格式：");
            lblMainBuf = AddKV(left, ref r, "缓冲 (ms)：");
            lblMainPass = AddKV(left, ref r, "直通/重采样：");
            lblMainPeriod = AddKV(left, ref r, "设备周期：");

            AddSeparator(left, ref r);
            AddHeader(left, ref r, "副输出（直播 / 推流）");
            lblAuxDev = AddKV(left, ref r, "设备：");
            lblAuxMode = AddKV(left, ref r, "模式：");
            lblAuxSync = AddKV(left, ref r, "同步：");
            lblAuxFmt = AddKV(left, ref r, "实际格式：");
            lblAuxBuf = AddKV(left, ref r, "缓冲 (ms)：");
            lblAuxPass = AddKV(left, ref r, "直通/重采样：");
            lblAuxPeriod = AddKV(left, ref r, "设备周期：");
            lblAuxQ = AddKV(left, ref r, "副通道重采样质量：");

            btnRefreshStatus = new Button { Text = "刷新状态", AutoSize = true, Dock = DockStyle.Top, Margin = new Padding(0, 8, 0, 0) };
            btnRefreshStatus.Click += (s, e) => FillStatus();
            left.Controls.Add(new Label(), 0, r); // 占位
            left.Controls.Add(btnRefreshStatus, 1, r++);
            left.RowCount = r;

            // ===== 右：用一个垂直 TableLayoutPanel 严格排序 =====
            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(10),
                AutoScroll = true
            };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            split.Panel2.Controls.Add(right);

            // 1) 设备组（右上）
            var gbDevices = MakeDeviceGroup();
            right.Controls.Add(gbDevices);
            right.SetRow(gbDevices, 0);

            // 2) 主输出组（中间）
            var gbMain = MakeMainGroup();
            right.Controls.Add(gbMain);
            right.SetRow(gbMain, 1);

            // 3) 副输出组（最下）
            var gbAux = MakeAuxGroup();
            right.Controls.Add(gbAux);
            right.SetRow(gbAux, 2);

            // 4) 全局/高级
            var gbAdv = MakeAdvancedGroup();
            right.Controls.Add(gbAdv);
            right.SetRow(gbAdv, 3);

            // 5) 按钮区
            var actions = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(0, 8, 0, 0) };
            btnOK = new Button { Text = "保存", AutoSize = true, Anchor = AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", AutoSize = true, Anchor = AnchorStyles.Right };
            btnOK.Click += (s, e) => OnSave();
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            // 按钮靠右
            var rightBtns = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            rightBtns.Controls.Add(btnOK);
            rightBtns.Controls.Add(btnCancel);
            actions.Controls.Add(rightBtns);
            right.Controls.Add(actions);
            right.SetRow(actions, 4);

            // 初始数据
            Result = null;
            FillDevices();
            ApplyCfgToUI(_orig);
            FillStatus();
        }

        // ===== 设备组 =====
        GroupBox MakeDeviceGroup()
        {
            var gb = new GroupBox { Text = "设备（输入 / 主输出 / 副输出）", Dock = DockStyle.Top, Padding = new Padding(8), Height = 138 };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            int r = 0;
            // 输入
            t.Controls.Add(new Label { Text = "输入：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(2, 10, 8, 0) }, 0, r);
            cmbIn = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 420, Margin = new Padding(0, 6, 8, 0) };
            t.Controls.Add(cmbIn, 1, r);
            btnRescan = new Button { Text = "刷新设备", AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(0, 6, 0, 0) };
            btnRescan.Click += (s, e) => FillDevices();
            t.Controls.Add(btnRescan, 2, r); r++;

            // 主输出
            t.Controls.Add(new Label { Text = "主输出：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(2, 10, 8, 0) }, 0, r);
            cmbMain = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 420, Margin = new Padding(0, 6, 8, 0) };
            t.Controls.Add(cmbMain, 1, r);
            t.Controls.Add(new Label(), 2, r); r++;

            // 副输出
            t.Controls.Add(new Label { Text = "副输出：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(2, 10, 8, 0) }, 0, r);
            cmbAux = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 420, Margin = new Padding(0, 6, 8, 0) };
            t.Controls.Add(cmbAux, 1, r);
            t.Controls.Add(new Label(), 2, r);

            gb.Controls.Add(t);
            return gb;
        }

        // ===== 主输出组 =====
        GroupBox MakeMainGroup()
        {
            var gb = new GroupBox { Text = "主输出（高音质 / 低延迟）", Dock = DockStyle.Top, Padding = new Padding(8), Height = 150 };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 3 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            int r = 0;
            t.Controls.Add(new Label { Text = "模式：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            cmbMainShare = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbMainShare.Items.AddRange(new object[] { "自动", "独占", "共享" });
            t.Controls.Add(cmbMainShare, 1, r);

            t.Controls.Add(new Label { Text = "同步：", AutoSize = true, Anchor = AnchorStyles.Left }, 2, r);
            cmbMainSync = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbMainSync.Items.AddRange(new object[] { "自动", "事件", "轮询" });
            t.Controls.Add(cmbMainSync, 3, r); r++;

            t.Controls.Add(new Label { Text = "采样率：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            cmbMainRate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbMainRate.Items.AddRange(new object[] { "44100", "48000", "88200", "96000", "192000" });
            t.Controls.Add(cmbMainRate, 1, r);

            t.Controls.Add(new Label { Text = "位深：", AutoSize = true, Anchor = AnchorStyles.Left }, 2, r);
            cmbMainBits = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbMainBits.Items.AddRange(new object[] { "16", "24", "32" });
            t.Controls.Add(cmbMainBits, 3, r); r++;

            t.Controls.Add(new Label { Text = "缓冲 (ms)：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            nudMainBuf = new NumericUpDown { Minimum = 4, Maximum = 500, Increment = 1, Width = 160, Anchor = AnchorStyles.Left };
            t.Controls.Add(nudMainBuf, 1, r);

            gb.Controls.Add(t);
            return gb;
        }

        // ===== 副输出组 =====
        GroupBox MakeAuxGroup()
        {
            var gb = new GroupBox { Text = "副输出（直播 / 推流）", Dock = DockStyle.Top, Padding = new Padding(8), Height = 190 };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 4 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            int r = 0;
            t.Controls.Add(new Label { Text = "模式：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            cmbAuxShare = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbAuxShare.Items.AddRange(new object[] { "自动", "独占", "共享" });
            t.Controls.Add(cmbAuxShare, 1, r);

            t.Controls.Add(new Label { Text = "同步：", AutoSize = true, Anchor = AnchorStyles.Left }, 2, r);
            cmbAuxSync = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbAuxSync.Items.AddRange(new object[] { "自动", "事件", "轮询" });
            t.Controls.Add(cmbAuxSync, 3, r); r++;

            t.Controls.Add(new Label { Text = "采样率：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            cmbAuxRate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbAuxRate.Items.AddRange(new object[] { "44100", "48000", "88200", "96000", "192000" });
            t.Controls.Add(cmbAuxRate, 1, r);

            t.Controls.Add(new Label { Text = "位深：", AutoSize = true, Anchor = AnchorStyles.Left }, 2, r);
            cmbAuxBits = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbAuxBits.Items.AddRange(new object[] { "16", "24", "32" });
            t.Controls.Add(cmbAuxBits, 3, r); r++;

            t.Controls.Add(new Label { Text = "缓冲 (ms)：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
            nudAuxBuf = new NumericUpDown { Minimum = 20, Maximum = 1000, Increment = 5, Width = 160, Anchor = AnchorStyles.Left };
            t.Controls.Add(nudAuxBuf, 1, r);

            t.Controls.Add(new Label { Text = "重采样质量：", AutoSize = true, Anchor = AnchorStyles.Left }, 2, r);
            cmbAuxQuality = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160, Anchor = AnchorStyles.Left };
            cmbAuxQuality.Items.AddRange(new object[] { "30", "40", "50" });
            t.Controls.Add(cmbAuxQuality, 3, r);

            gb.Controls.Add(t);
            return gb;
        }

        // ===== 全局/高级组 =====
        GroupBox MakeAdvancedGroup()
        {
            var gb = new GroupBox { Text = "全局 / 高级", Dock = DockStyle.Top, Padding = new Padding(8), Height = 90 };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 2 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            chkAutoStart = new CheckBox { Text = "开机自启动", AutoSize = true, Anchor = AnchorStyles.Left };
            chkLogging = new CheckBox { Text = "启用日志（排障时开启）", AutoSize = true, Anchor = AnchorStyles.Left };
            chkForceUserFormat = new CheckBox { Text = "强制使用右侧格式（仅独占）", AutoSize = true, Anchor = AnchorStyles.Left };
            chkForcePassthroughStrict = new CheckBox { Text = "强制直通（不可直通则不播放）", AutoSize = true, Anchor = AnchorStyles.Left };

            t.Controls.Add(chkAutoStart, 0, 0);
            t.Controls.Add(chkLogging, 1, 0);
            t.Controls.Add(chkForceUserFormat, 2, 0);
            t.Controls.Add(chkForcePassthroughStrict, 3, 0);

            gb.Controls.Add(t);
            return gb;
        }

        // ===== 状态区辅助 =====
        static void AddHeader(TableLayoutPanel t, ref int r, string text)
        {
            var lbl = new Label { Text = text, AutoSize = true, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold), Margin = new Padding(0, 4, 0, 6) };
            t.Controls.Add(lbl, 0, r);
            t.SetColumnSpan(lbl, 2);
            r++;
        }

        static void AddSeparator(TableLayoutPanel t, ref int r)
        {
            var line = new Label { BorderStyle = BorderStyle.Fixed3D, Height = 2, Dock = DockStyle.Top, Margin = new Padding(0, 6, 0, 6) };
            t.Controls.Add(line, 0, r);
            t.SetColumnSpan(line, 2);
            r++;
        }

        static Label AddKV(TableLayoutPanel t, ref int r, string key)
        {
            t.Controls.Add(new Label { Text = key, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 4, 8, 0) }, 0, r);
            var v = new Label { Text = "-", AutoSize = true, MaximumSize = new Size(500, 0), Margin = new Padding(0, 4, 0, 0) };
            t.Controls.Add(v, 1, r);
            r++;
            return v;
        }

        // ===== 设备枚举/填充 =====
        void FillDevices()
        {
            string selIn = GetSelectedId(cmbIn);
            string selM = GetSelectedId(cmbMain);
            string selA = GetSelectedId(cmbAux);

            FillDeviceCombo(cmbIn, includeCapture: true, includeRender: true, preferId: _orig.InputDeviceId);
            FillDeviceCombo(cmbMain, includeCapture: false, includeRender: true, preferId: _orig.MainDeviceId);
            FillDeviceCombo(cmbAux, includeCapture: false, includeRender: true, preferId: _orig.AuxDeviceId);

            // 恢复选择（优先 id）
            SelectById(cmbIn, string.IsNullOrEmpty(_orig.InputDeviceId) ? selIn : _orig.InputDeviceId);
            SelectById(cmbMain, string.IsNullOrEmpty(_orig.MainDeviceId) ? selM : _orig.MainDeviceId);
            SelectById(cmbAux, string.IsNullOrEmpty(_orig.AuxDeviceId) ? selA : _orig.AuxDeviceId);
        }

        void FillDeviceCombo(ComboBox cmb, bool includeCapture, bool includeRender, string preferId)
        {
            cmb.Items.Clear();
            var list = new List<ListItem>();

            if (includeCapture)
            {
                foreach (var d in _mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                    list.Add(new ListItem(d.FriendlyName, d.ID));
                // 输入允许选择环回（Render 作为输入）
                foreach (var d in _mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                    list.Add(new ListItem(d.FriendlyName + "（环回）", d.ID));
            }
            if (includeRender)
            {
                foreach (var d in _mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                    list.Add(new ListItem(d.FriendlyName, d.ID));
            }

            foreach (var it in list) cmb.Items.Add(it);

            if (!string.IsNullOrEmpty(preferId) && !SelectById(cmb, preferId))
            {
                if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
            }
            else
            {
                if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
            }
        }

        // ===== 状态刷新 =====
        void FillStatus()
        {
            try
            {
                var s = _getStatus != null ? _getStatus() : null;
                if (s == null)
                {
                    SetLabel(lblRunning, "—"); SetLabel(lblInDev, "—"); SetLabel(lblInRole, "—"); SetLabel(lblInFmt, "—");
                    SetLabel(lblMainDev, "—"); SetLabel(lblMainMode, "—"); SetLabel(lblMainSync, "—"); SetLabel(lblMainFmt, "—");
                    SetLabel(lblMainBuf, "—"); SetLabel(lblMainPass, "—"); SetLabel(lblMainPeriod, "—");
                    SetLabel(lblAuxDev, "—"); SetLabel(lblAuxMode, "—"); SetLabel(lblAuxSync, "—"); SetLabel(lblAuxFmt, "—");
                    SetLabel(lblAuxBuf, "—"); SetLabel(lblAuxPass, "—"); SetLabel(lblAuxPeriod, "—"); SetLabel(lblAuxQ, "—");
                    return;
                }

                SetLabel(lblRunning, s.Running ? "是" : "否");
                SetLabel(lblInDev, s.InputDevice);
                SetLabel(lblInRole, s.InputRole);
                SetLabel(lblInFmt, s.InputFormat);

                SetLabel(lblMainDev, s.MainDevice);
                SetLabel(lblMainMode, s.MainMode);
                SetLabel(lblMainSync, s.MainSync);
                SetLabel(lblMainFmt, s.MainFormat);
                SetLabel(lblMainBuf, s.MainBufferMs.ToString());
                SetLabel(lblMainPass, s.MainPassDesc);
                SetLabel(lblMainPeriod, string.Format("默认 {0:0.###} ms / 最小 {1:0.###} ms", s.MainDefaultPeriodMs, s.MainMinimumPeriodMs));

                SetLabel(lblAuxDev, s.AuxDevice);
                SetLabel(lblAuxMode, s.AuxMode);
                SetLabel(lblAuxSync, s.AuxSync);
                SetLabel(lblAuxFmt, s.AuxFormat);
                SetLabel(lblAuxBuf, s.AuxBufferMs.ToString());
                SetLabel(lblAuxPass, s.AuxPassDesc);
                SetLabel(lblAuxPeriod, string.Format("默认 {0:0.###} ms / 最小 {1:0.###} ms", s.AuxDefaultPeriodMs, s.AuxMinimumPeriodMs));
                SetLabel(lblAuxQ, s.AuxQuality.ToString());
            }
            catch
            {
                // 吞掉 UI 刷新异常，避免影响使用
            }
        }

        static void SetLabel(Label l, string text)
        {
            if (l == null) return;
            l.Text = string.IsNullOrEmpty(text) ? "-" : text;
        }

        // ===== 配置 <-> UI =====
        void ApplyCfgToUI(AppSettings c)
        {
            SelectById(cmbIn, c.InputDeviceId);
            SelectById(cmbMain, c.MainDeviceId);
            SelectById(cmbAux, c.AuxDeviceId);

            cmbMainShare.SelectedIndex = ToIndex(c.MainShare);
            cmbMainSync.SelectedIndex = ToIndex(c.MainSync);
            SetComboValue(cmbMainRate, c.MainRate);
            SetComboValue(cmbMainBits, c.MainBits);
            nudMainBuf.Value = FixRange(c.MainBufMs, (int)nudMainBuf.Minimum, (int)nudMainBuf.Maximum);

            cmbAuxShare.SelectedIndex = ToIndex(c.AuxShare);
            cmbAuxSync.SelectedIndex = ToIndex(c.AuxSync);
            SetComboValue(cmbAuxRate, c.AuxRate);
            SetComboValue(cmbAuxBits, c.AuxBits);
            nudAuxBuf.Value = FixRange(c.AuxBufMs, (int)nudAuxBuf.Minimum, (int)nudAuxBuf.Maximum);
            SetComboValue(cmbAuxQuality, c.AuxResamplerQuality);

            chkAutoStart.Checked = c.AutoStart;
            chkLogging.Checked = c.EnableLogging;
            chkForceUserFormat.Checked = c.ForceUserFormat;
            chkForcePassthroughStrict.Checked = c.ForcePassthroughStrict;
        }

        void OnSave()
        {
            var c = new AppSettings();
            c.InputDeviceId = GetSelectedId(cmbIn);
            c.MainDeviceId = GetSelectedId(cmbMain);
            c.AuxDeviceId = GetSelectedId(cmbAux);

            c.MainShare = FromIndexShare(cmbMainShare.SelectedIndex);
            c.MainSync = FromIndexSync(cmbMainSync.SelectedIndex);
            c.MainRate = ParseComboInt(cmbMainRate, 192000);
            c.MainBits = ParseComboInt(cmbMainBits, 24);
            c.MainBufMs = (int)nudMainBuf.Value;

            c.AuxShare = FromIndexShare(cmbAuxShare.SelectedIndex);
            c.AuxSync = FromIndexSync(cmbAuxSync.SelectedIndex);
            c.AuxRate = ParseComboInt(cmbAuxRate, 44100);
            c.AuxBits = ParseComboInt(cmbAuxBits, 16);
            c.AuxBufMs = (int)nudAuxBuf.Value;
            c.AuxResamplerQuality = ParseComboInt(cmbAuxQuality, 40);

            c.AutoStart = chkAutoStart.Checked;
            c.EnableLogging = chkLogging.Checked;
            c.ForceUserFormat = chkForceUserFormat.Checked;
            c.ForcePassthroughStrict = chkForcePassthroughStrict.Checked;

            Result = c;
            DialogResult = DialogResult.OK;
            Close();
        }

        // ===== 工具 =====
        static int FixRange(int v, int lo, int hi) { if (v < lo) return lo; if (v > hi) return hi; return v; }

        static void SetComboValue(ComboBox cmb, int v)
        {
            if (cmb == null) return;
            string s = v.ToString();
            int idx = -1;
            for (int i = 0; i < cmb.Items.Count; i++)
                if (string.Equals(cmb.Items[i].ToString(), s, StringComparison.OrdinalIgnoreCase)) { idx = i; break; }
            if (idx >= 0) cmb.SelectedIndex = idx;
            else { cmb.Items.Add(s); cmb.SelectedIndex = cmb.Items.Count - 1; }
        }

        static int ParseComboInt(ComboBox cmb, int @default)
        {
            if (cmb == null || cmb.SelectedItem == null) return @default;
            int v; if (int.TryParse(cmb.SelectedItem.ToString(), out v)) return v;
            return @default;
        }

        static bool SelectById(ComboBox cmb, string id)
        {
            if (cmb == null || string.IsNullOrEmpty(id)) return false;
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                var li = cmb.Items[i] as ListItem;
                if (li != null && string.Equals(li.Id, id, StringComparison.OrdinalIgnoreCase))
                { cmb.SelectedIndex = i; return true; }
            }
            return false;
        }
        static string GetSelectedId(ComboBox cmb)
        {
            if (cmb == null || cmb.SelectedItem == null) return null;
            var li = cmb.SelectedItem as ListItem;
            return li != null ? li.Id : null;
        }

        static ShareModeOption FromIndexShare(int i)
        {
            if (i == 1) return ShareModeOption.Exclusive;
            if (i == 2) return ShareModeOption.Shared;
            return ShareModeOption.Auto;
        }
        static SyncModeOption FromIndexSync(int i)
        {
            if (i == 1) return SyncModeOption.Event;
            if (i == 2) return SyncModeOption.Polling;
            return SyncModeOption.Auto;
        }
        static int ToIndex(ShareModeOption s)
        {
            if (s == ShareModeOption.Exclusive) return 1;
            if (s == ShareModeOption.Shared) return 2;
            return 0;
        }
        static int ToIndex(SyncModeOption s)
        {
            if (s == SyncModeOption.Event) return 1;
            if (s == SyncModeOption.Polling) return 2;
            return 0;
        }

        sealed class ListItem
        {
            public readonly string Text;
            public readonly string Id;
            public ListItem(string text, string id) { Text = text; Id = id; }
            public override string ToString() { return Text; }
        }
    }
}
