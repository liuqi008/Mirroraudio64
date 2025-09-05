using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    internal sealed class SettingsForm : Form
    {
        private readonly Func<StatusSnapshot> _getStatus;
        public AppSettings Result { get; private set; }

        private ComboBox cbInput, cbMain, cbAux;
        private ComboBox cbMainShare, cbMainSync, cbAuxShare, cbAuxSync, cbMainBits, cbAuxBits, cbMainRate, cbAuxRate;
        private NumericUpDown nudMainBuf, nudAuxBuf;
        private CheckBox chkAutoStart, chkLog;

        private Label lbInput, lbMain, lbAux, lbMainMode, lbAuxMode, lbMainFmt, lbAuxFmt, lbMainBuf, lbAuxBuf, lbMainPer, lbAuxPer;

        public SettingsForm(AppSettings current, Func<StatusSnapshot> getStatus)
        {
            Result = Clone(current);
            _getStatus = getStatus;

            Text = "MirrorAudio 设置";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(680, 520);
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;

            BuildUi();
            LoadValues(current);
            RefreshStatus();

            var timer = new Timer { Interval = 800 };
            timer.Tick += (s, e) => { if (!IsDisposed) RefreshStatus(); };
            timer.Start();
            FormClosed += (s, e) => timer.Stop();
        }

        private static AppSettings Clone(AppSettings s) =>
            new AppSettings
            {
                InputDeviceId = s.InputDeviceId,
                MainDeviceId = s.MainDeviceId,
                AuxDeviceId = s.AuxDeviceId,
                MainShare = s.MainShare,
                MainSync = s.MainSync,
                MainRate = s.MainRate,
                MainBits = s.MainBits,
                MainBufMs = s.MainBufMs,
                AuxShare = s.AuxShare,
                AuxSync = s.AuxSync,
                AuxRate = s.AuxRate,
                AuxBits = s.AuxBits,
                AuxBufMs = s.AuxBufMs,
                AutoStart = s.AutoStart,
                EnableLogging = s.EnableLogging
            };

        private sealed class DevItem
        {
            public string Id { get; }
            public string Name { get; }
            public DevItem(string id, string name) { Id = id; Name = name; }
            public override string ToString() => Name;
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(12),
                AutoSize = true
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            Controls.Add(root);

            // 左：配置区
            var left = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true
            };
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(left, 0, 0);

            // 右：状态区
            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true
            };
            root.Controls.Add(right, 1, 0);

            // —— 设备选择 —— //
            left.Controls.Add(new Label { Text = "输入源：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
            cbInput = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            left.Controls.Add(cbInput, 1, 0);

            left.Controls.Add(new Label { Text = "主输出设备：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 1);
            cbMain = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            left.Controls.Add(cbMain, 1, 1);

            left.Controls.Add(new Label { Text = "副输出设备：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 2);
            cbAux = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            left.Controls.Add(cbAux, 1, 2);

            left.Controls.Add(new Label { Text = "—— 主通道 ——", AutoSize = true, ForeColor = SystemColors.GrayText }, 0, 3);
            left.SetColumnSpan(left.Controls[left.Controls.Count - 1], 2);

            left.Controls.Add(new Label { Text = "共享/独占：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 4);
            cbMainShare = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbMainShare.Items.AddRange(new object[] { ShareModeOption.Auto, ShareModeOption.Exclusive, ShareModeOption.Shared });
            left.Controls.Add(cbMainShare, 1, 4);

            left.Controls.Add(new Label { Text = "同步模式：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 5);
            cbMainSync = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbMainSync.Items.AddRange(new object[] { SyncModeOption.Auto, SyncModeOption.Event, SyncModeOption.Polling });
            left.Controls.Add(cbMainSync, 1, 5);

            left.Controls.Add(new Label { Text = "采样率(独占)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 6);
            cbMainRate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbMainRate.Items.AddRange(new object[] { 44100, 48000, 88200, 96000, 176400, 192000 });
            left.Controls.Add(cbMainRate, 1, 6);

            left.Controls.Add(new Label { Text = "位深(独占)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 7);
            cbMainBits = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbMainBits.Items.AddRange(new object[] { 16, 24, 32 });
            left.Controls.Add(cbMainBits, 1, 7);

            left.Controls.Add(new Label { Text = "缓冲(ms)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 8);
            nudMainBuf = new NumericUpDown { Minimum = 2, Maximum = 1000, Increment = 1, Anchor = AnchorStyles.Left, Width = 100 };
            left.Controls.Add(nudMainBuf, 1, 8);

            left.Controls.Add(new Label { Text = "—— 副通道 ——", AutoSize = true, ForeColor = SystemColors.GrayText }, 0, 9);
            left.SetColumnSpan(left.Controls[left.Controls.Count - 1], 2);

            left.Controls.Add(new Label { Text = "共享/独占：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 10);
            cbAuxShare = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbAuxShare.Items.AddRange(new object[] { ShareModeOption.Shared, ShareModeOption.Auto, ShareModeOption.Exclusive });
            left.Controls.Add(cbAuxShare, 1, 10);

            left.Controls.Add(new Label { Text = "同步模式：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 11);
            cbAuxSync = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbAuxSync.Items.AddRange(new object[] { SyncModeOption.Auto, SyncModeOption.Event, SyncModeOption.Polling });
            left.Controls.Add(cbAuxSync, 1, 11);

            left.Controls.Add(new Label { Text = "采样率(独占)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 12);
            cbAuxRate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbAuxRate.Items.AddRange(new object[] { 44100, 48000, 88200, 96000, 176400, 192000 });
            left.Controls.Add(cbAuxRate, 1, 12);

            left.Controls.Add(new Label { Text = "位深(独占)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 13);
            cbAuxBits = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left };
            cbAuxBits.Items.AddRange(new object[] { 16, 24, 32 });
            left.Controls.Add(cbAuxBits, 1, 13);

            left.Controls.Add(new Label { Text = "缓冲(ms)：", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 14);
            nudAuxBuf = new NumericUpDown { Minimum = 10, Maximum = 2000, Increment = 5, Anchor = AnchorStyles.Left, Width = 100 };
            left.Controls.Add(nudAuxBuf, 1, 14);

            // 其他
            chkAutoStart = new CheckBox { Text = "开机自启动", AutoSize = true, Anchor = AnchorStyles.Left };
            chkLog = new CheckBox { Text = "启用日志（便于排障）", AutoSize = true, Anchor = AnchorStyles.Left };
            left.Controls.Add(chkAutoStart, 0, 15);
            left.Controls.Add(chkLog, 1, 15);

            // 按钮
            var pnlBtn = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Height = 40 };
            var btnOK = new Button { Text = "确定", DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel };
            pnlBtn.Controls.Add(btnOK);
            pnlBtn.Controls.Add(btnCancel);
            Controls.Add(pnlBtn);

            btnOK.Click += (s, e) => { SaveValues(); };
            btnCancel.Click += (s, e) => { /* no-op */ };

            // —— 右侧状态区 —— //
            var gbStatus = new GroupBox { Text = "当前运行状态", Dock = DockStyle.Fill };
            right.Controls.Add(gbStatus);

            var st = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, Padding = new Padding(8) };
            st.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            st.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            gbStatus.Controls.Add(st);

            lbInput = AddRow(st, "输入：");
            lbMain  = AddRow(st, "主输出：");
            lbAux   = AddRow(st, "副输出：");
            lbMainMode = AddRow(st, "主通道模式：");
            lbAuxMode  = AddRow(st, "副通道模式：");
            lbMainFmt  = AddRow(st, "主通道格式：");
            lbAuxFmt   = AddRow(st, "副通道格式：");
            lbMainBuf  = AddRow(st, "主缓冲(ms)：");
            lbAuxBuf   = AddRow(st, "副缓冲(ms)：");
            lbMainPer  = AddRow(st, "主设备周期(ms)：");
            lbAuxPer   = AddRow(st, "副设备周期(ms)：");
        }

        private static Label AddRow(TableLayoutPanel tl, string name)
        {
            tl.Controls.Add(new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left });
            var v = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left };
            tl.Controls.Add(v);
            return v;
        }

        private void LoadValues(AppSettings c)
        {
            // 设备列表
            cbInput.Items.Clear();
            cbMain.Items.Clear();
            cbAux.Items.Clear();

            using (var mm = new MMDeviceEnumerator())
            {
                // 输入（既包含录音也包含渲染，便于“环回”选择）
                foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                    cbInput.Items.Add(new DevItem(d.ID, "[录音] " + d.FriendlyName));
                foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                    cbInput.Items.Add(new DevItem(d.ID, "[环回] " + d.FriendlyName));

                foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                {
                    cbMain.Items.Add(new DevItem(d.ID, d.FriendlyName));
                    cbAux.Items.Add(new DevItem(d.ID, d.FriendlyName));
                }
            }

            SelectById(cbInput, c.InputDeviceId);
            SelectById(cbMain,  c.MainDeviceId);
            SelectById(cbAux,   c.AuxDeviceId);

            cbMainShare.SelectedItem = c.MainShare;
            cbMainSync.SelectedItem  = c.MainSync;
            cbMainRate.SelectedItem  = (object)c.MainRate ?? 192000;
            cbMainBits.SelectedItem  = (object)c.MainBits ?? 24;
            nudMainBuf.Value         = Math.Max(2, Math.Min(1000, c.MainBufMs));

            cbAuxShare.SelectedItem  = c.AuxShare;
            cbAuxSync.SelectedItem   = c.AuxSync;
            cbAuxRate.SelectedItem   = (object)c.AuxRate ?? 48000;
            cbAuxBits.SelectedItem   = (object)c.AuxBits ?? 16;
            nudAuxBuf.Value          = Math.Max(10, Math.Min(2000, c.AuxBufMs));

            chkAutoStart.Checked = c.AutoStart;
            chkLog.Checked = c.EnableLogging;
        }

        private static void SelectById(ComboBox cb, string id)
        {
            if (string.IsNullOrEmpty(id) || cb.Items.Count == 0) { if (cb.Items.Count > 0) cb.SelectedIndex = 0; return; }
            for (int i = 0; i < cb.Items.Count; i++)
            {
                var it = cb.Items[i] as DevItem;
                if (it != null && string.Equals(it.Id, id, StringComparison.OrdinalIgnoreCase))
                { cb.SelectedIndex = i; return; }
            }
            cb.SelectedIndex = 0;
        }

        private void SaveValues()
        {
            Result.InputDeviceId = (cbInput.SelectedItem as DevItem)?.Id;
            Result.MainDeviceId  = (cbMain .SelectedItem as DevItem)?.Id;
            Result.AuxDeviceId   = (cbAux  .SelectedItem as DevItem)?.Id;

            Result.MainShare = (ShareModeOption)cbMainShare.SelectedItem;
            Result.MainSync  = (SyncModeOption)cbMainSync.SelectedItem;
            Result.MainRate  = (int)cbMainRate.SelectedItem;
            Result.MainBits  = (int)cbMainBits.SelectedItem;
            Result.MainBufMs = (int)nudMainBuf.Value;

            Result.AuxShare  = (ShareModeOption)cbAuxShare.SelectedItem;
            Result.AuxSync   = (SyncModeOption)cbAuxSync.SelectedItem;
            Result.AuxRate   = (int)cbAuxRate.SelectedItem;
            Result.AuxBits   = (int)cbAuxBits.SelectedItem;
            Result.AuxBufMs  = (int)nudAuxBuf.Value;

            Result.AutoStart = chkAutoStart.Checked;
            Result.EnableLogging = chkLog.Checked;
        }

        private void RefreshStatus()
        {
            StatusSnapshot s;
            try { s = _getStatus?.Invoke(); }
            catch { s = null; }
            if (s == null) return;

            lbInput.Text   = $"{s.InputRole} | {s.InputFormat} | {s.InputDevice}";
            lbMain.Text    = s.MainDevice;
            lbAux.Text     = s.AuxDevice;

            lbMainMode.Text = $"{s.MainMode}/{s.MainSync}";
            lbAuxMode.Text  = $"{s.AuxMode}/{s.AuxSync}";

            lbMainFmt.Text  = s.MainFormat;
            lbAuxFmt.Text   = s.AuxFormat;

            lbMainBuf.Text  = s.MainBufferMs.ToString();
            lbAuxBuf.Text   = s.AuxBufferMs.ToString();

            lbMainPer.Text  = $"{s.MainDefaultPeriodMs:0.###} / {s.MainMinimumPeriodMs:0.###}";
            lbAuxPer.Text   = $"{s.AuxDefaultPeriodMs:0.###} / {s.AuxMinimumPeriodMs:0.###}";
        }
    }
}
