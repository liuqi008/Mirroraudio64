
using System;
using System.Drawing;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    sealed class SettingsForm : Form
    {
        readonly Func<StatusSnapshot> _statusProvider;

        // 左侧状态
        readonly Label lblRun=new Label(), lblMain=new Label(), lblAux=new Label(),
                       lblMainFmt=new Label(), lblAuxFmt=new Label(), lblMainBuf=new Label(), lblAuxBuf=new Label(),
                       lblMainPer=new Label(), lblAuxPer=new Label(),
                       lblMainRaw=new Label(), lblAuxRaw=new Label();

        // 右侧设置
        readonly ComboBox cmbMainDev=new ComboBox(), cmbAuxDev=new ComboBox();
        readonly ComboBox cmbMainShare=new ComboBox(), cmbAuxShare=new ComboBox();
        readonly ComboBox cmbMainSync=new ComboBox(),  cmbAuxSync=new ComboBox();
        readonly NumericUpDown numMainBuf=new NumericUpDown(), numAuxBuf=new NumericUpDown();
        readonly NumericUpDown numMainRate=new NumericUpDown(), numAuxRate=new NumericUpDown();
        readonly ComboBox cmbMainBits=new ComboBox(), cmbAuxBits=new ComboBox();

        readonly Button btnOK=new Button(), btnCancel=new Button(), btnRefresh=new Button();
        readonly Timer tStatus = new Timer();

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings cur, Func<StatusSnapshot> statusProvider)
        {
            _statusProvider = statusProvider ?? (() => new StatusSnapshot { Running=false });

            Text = "MirrorAudio 设置";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = SystemFonts.MessageBoxFont;
            MinimumSize = new Size(980, 620);

            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 520 };
            Controls.Add(split);

            // 左侧：状态
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoScroll = true };
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            split.Panel1.Controls.Add(left);

            AddRow(left, "运行状态", lblRun);
            AddRow(left, "主输出", lblMain);
            AddRow(left, "副输出", lblAux);
            AddRow(left, "主格式", lblMainFmt);
            AddRow(left, "副格式", lblAuxFmt);
            AddRow(left, "主缓冲", lblMainBuf);
            AddRow(left, "副缓冲", lblAuxBuf);
            AddRow(left, "主周期", lblMainPer);
            AddRow(left, "副周期", lblAuxPer);
            AddRow(left, "主直通(RAW)", lblMainRaw);
            AddRow(left, "副直通(RAW)", lblAuxRaw);

            // 右侧：设置
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoScroll = true };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            split.Panel2.Controls.Add(right);

            AddRow(right, "主设备", cmbMainDev);
            AddRow(right, "副设备", cmbAuxDev);
            AddRow(right, "主共享模式", cmbMainShare);
            AddRow(right, "副共享模式", cmbAuxShare);
            AddRow(right, "主同步", cmbMainSync);
            AddRow(right, "副同步", cmbAuxSync);
            AddRow(right, "主缓冲(ms)", numMainBuf);
            AddRow(right, "副缓冲(ms)", numAuxBuf);
            AddRow(right, "主采样率", numMainRate);
            AddRow(right, "副采样率", numAuxRate);
            AddRow(right, "主位深", cmbMainBits);
            AddRow(right, "副位深", cmbAuxBits);

            var btns = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 48 };
            btnOK.Text = "确定"; btnCancel.Text = "取消"; btnRefresh.Text = "刷新设备";
            btnOK.Click += (s,e)=>{ Result = Collect(); DialogResult = DialogResult.OK; };
            btnCancel.Click += (s,e)=>{ DialogResult = DialogResult.Cancel; };
            btnRefresh.Click += (s,e)=>PopulateDevices();
            btns.Controls.AddRange(new Control[]{ btnOK, btnCancel, btnRefresh });
            split.Panel2.Controls.Add(btns);

            foreach (var cb in new[]{ cmbMainShare, cmbAuxShare })
                cb.Items.AddRange(new object[]{ ShareModeOption.Auto, ShareModeOption.Exclusive, ShareModeOption.Shared });
            foreach (var cb in new[]{ cmbMainSync, cmbAuxSync })
                cb.Items.AddRange(new object[]{ SyncModeOption.Auto, SyncModeOption.Event, SyncModeOption.Polling });
            foreach (var cb in new[]{ cmbMainBits, cmbAuxBits })
                cb.Items.AddRange(new object[]{ 16, 24, 32 });

            foreach (var n in new[]{ numMainBuf, numAuxBuf }) { n.Minimum = 2; n.Maximum = 200; n.Value = 6; n.DecimalPlaces = 0; n.Increment = 1; }
            foreach (var n in new[]{ numMainRate, numAuxRate }) { n.Minimum = 8000; n.Maximum = 384000; n.Value = 48000; n.DecimalPlaces = 0; n.Increment = 1000; }

            // load current
            Apply(cur);
            PopulateDevices();

            tStatus.Interval = 500;
            tStatus.Tick += (s,e)=>RenderStatus();
            tStatus.Start();
        }

        void AddRow(TableLayoutPanel tl, string name, Control value)
        {
            tl.RowCount += 1;
            tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var l = new Label { Text = name, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill, AutoSize = true };
            value.Dock = DockStyle.Fill;
            tl.Controls.Add(l, 0, tl.RowCount - 1);
            tl.Controls.Add(value, 1, tl.RowCount - 1);
        }

        void PopulateDevices()
        {
            try
            {
                var de = new MMDeviceEnumerator();
                var list = de.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
                Populate(cmbMainDev, list);
                Populate(cmbAuxDev, list);
            }
            catch {}
        }

        void Populate(ComboBox cb, MMDeviceCollection col)
        {
            var sel = cb.SelectedItem as DevItem;
            cb.Items.Clear();
            foreach (var d in col)
            {
                cb.Items.Add(new DevItem { Id = d.ID, Name = d.FriendlyName });
            }
            if (sel != null)
            {
                foreach (var o in cb.Items)
                {
                    var di = o as DevItem;
                    if (di != null && di.Id == sel.Id) { cb.SelectedItem = o; break; }
                }
            }
            if (cb.SelectedIndex < 0 && cb.Items.Count > 0) cb.SelectedIndex = 0;
        }

        void RenderStatus()
        {
            StatusSnapshot s = null;
            try { s = _statusProvider(); } catch {}
            if (s == null) return;
            lblRun.Text = s.Running ? "运行中" : "停止";
            lblMain.Text = (s.MainDevice ?? "-") + " | " + (s.MainMode ?? "-") + " | " + (s.MainSync ?? "-");
            lblAux.Text  = (s.AuxDevice  ?? "-") + " | " + (s.AuxMode  ?? "-") + " | " + (s.AuxSync  ?? "-");
            lblMainFmt.Text = s.MainFormat ?? "-";
            lblAuxFmt.Text  = s.AuxFormat  ?? "-";
            lblMainBuf.Text = s.MainBufferMs > 0 ? (s.MainBufferMs + " ms") : "-";
            lblAuxBuf.Text  = s.AuxBufferMs  > 0 ? (s.AuxBufferMs  + " ms") : "-";
            lblMainPer.Text = string.Format("默认 {0:0.##} ms / 最小 {1:0.##} ms", s.MainDefaultPeriodMs, s.MainMinimumPeriodMs);
            lblAuxPer.Text  = string.Format("默认 {0:0.##} ms / 最小 {1:0.##} ms", s.AuxDefaultPeriodMs, s.AuxMinimumPeriodMs);
            lblMainRaw.Text = s.MainMode == "独占" ? (s.MainRaw ? "已启用" : "未启用") : "-";
            lblAuxRaw.Text  = s.AuxMode  == "独占" ? (s.AuxRaw  ? "已启用" : "未启用") : "-";
        }

        void Apply(AppSettings a)
        {
            if (a == null) a = new AppSettings();
            cmbMainShare.SelectedItem = a.MainShare;
            cmbAuxShare.SelectedItem  = a.AuxShare;
            cmbMainSync.SelectedItem  = a.MainSync;
            cmbAuxSync.SelectedItem   = a.AuxSync;
            numMainBuf.Value = Clamp(numMainBuf, a.MainBufMs);
            numAuxBuf.Value  = Clamp(numAuxBuf, a.AuxBufMs);
            numMainRate.Value = Clamp(numMainRate, a.MainRate);
            numAuxRate.Value  = Clamp(numAuxRate, a.AuxRate);
            cmbMainBits.SelectedItem = a.MainBits;
            cmbAuxBits.SelectedItem  = a.AuxBits;
        }

        decimal Clamp(NumericUpDown n, int v)
        {
            if (v < n.Minimum) v = (int)n.Minimum;
            if (v > n.Maximum) v = (int)n.Maximum;
            return v;
        }

        AppSettings Collect()
        {
            var a = new AppSettings();
            var md = cmbMainDev.SelectedItem as DevItem;
            var ad = cmbAuxDev .SelectedItem as DevItem;
            a.MainDeviceId = md != null ? md.Id : null;
            a.AuxDeviceId  = ad != null ? ad.Id : null;
            a.MainShare = (ShareModeOption)(cmbMainShare.SelectedItem ?? ShareModeOption.Exclusive);
            a.AuxShare  = (ShareModeOption)(cmbAuxShare .SelectedItem ?? ShareModeOption.Exclusive);
            a.MainSync  = (SyncModeOption)(cmbMainSync.SelectedItem ?? SyncModeOption.Event);
            a.AuxSync   = (SyncModeOption)(cmbAuxSync .SelectedItem ?? SyncModeOption.Event);
            a.MainBufMs = (int)numMainBuf.Value;
            a.AuxBufMs  = (int)numAuxBuf .Value;
            a.MainRate  = (int)numMainRate.Value;
            a.AuxRate   = (int)numAuxRate .Value;
            a.MainBits  = (int)(cmbMainBits.SelectedItem ?? 24);
            a.AuxBits   = (int)(cmbAuxBits .SelectedItem ?? 16);
            return a;
        }

        sealed class DevItem { public string Id; public string Name; public override string ToString(){ return Name; } }
    }
}
