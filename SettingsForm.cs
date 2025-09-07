using System;
using System.Drawing;
using System.Windows.Forms;

namespace MirrorAudio
{
    public partial class SettingsForm : Form
    {
        private readonly Func<Program.StatusSnapshot> _getStatus;
        private readonly Action<Program.AppSettings> _saveConfig;
        private readonly Func<Program.AppSettings> _getConfig;
        private readonly Action _applyAndRestart;
        private readonly Action _cleanup;

        // Controls (subset)
        CheckBox chkMainExclusive, chkAuxExclusive;
        CheckBox chkMainEvent, chkAuxEvent;
        NumericUpDown numMainRate, numMainBits, numMainBuf;
        NumericUpDown numAuxRate, numAuxBits, numAuxBuf;

        // New controls
        CheckBox chkMainStrict, chkAuxStrict;
        NumericUpDown numMainAlignN, numAuxAlignN;

        // Status labels (left)
        Label lblMainMap, lblAuxMap;
        Label lblMainFmt, lblAuxFmt;
        Label lblMainMode, lblAuxMode;
        Label lblMainBuf, lblAuxBuf;
        Label lblMainRes, lblAuxRes;

        Button btnApply, btnSaveExit;

        public SettingsForm(
            Func<Program.StatusSnapshot> getStatus,
            Action<Program.AppSettings> saveConfig,
            Func<Program.AppSettings> getConfig,
            Action applyAndRestart,
            Action cleanup)
        {
            _getStatus = getStatus;
            _saveConfig = saveConfig;
            _getConfig = getConfig;
            _applyAndRestart = applyAndRestart;
            _cleanup = cleanup;

            InitializeComponent();
            BuildUI();
            LoadConfig();
        }

        void InitializeComponent()
        {
            this.Text = "MirrorAudio 设置";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(980, 640);
            this.FormClosing += (_, e) => _cleanup?.Invoke();
        }

        void BuildUI()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            this.Controls.Add(root);

            // LEFT: Status
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 10, Padding = new Padding(12) };
            left.RowStyles.Clear();
            for (int i = 0; i < 10; i++) left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(left, 0, 0);

            lblMainMode = AddKV(left, "主通道模式：");
            lblMainFmt  = AddKV(left, "主通道格式：");
            lblMainMap  = AddKV(left, "主位深（容器→线缆）：");
            lblMainBuf  = AddKV(left, "主缓冲(ms)：");
            lblMainRes  = AddKV(left, "主直通/重采样：");

            left.Controls.Add(new Label { Text = "—", AutoSize = true, Margin = new Padding(0, 8, 0, 8) });

            lblAuxMode = AddKV(left, "副通道模式：");
            lblAuxFmt  = AddKV(left, "副通道格式：");
            lblAuxMap  = AddKV(left, "副位深（容器→线缆）：");
            lblAuxBuf  = AddKV(left, "副缓冲(ms)：");
            lblAuxRes  = AddKV(left, "副直通/重采样：");

            // RIGHT: Settings
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 12, Padding = new Padding(12) };
            for (int i = 0; i < 12; i++) right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 56));
            root.Controls.Add(right, 1, 0);

            // Main group
            AddHeader(right, "主输出");
            chkMainExclusive = AddCheck(right, "独占模式");
            chkMainEvent = AddCheck(right, "事件驱动");
            numMainRate = AddNum(right, "采样率(Hz)", 8000, 384000, 192000, 100);
            numMainBits = AddNum(right, "位深(bit)", 16, 32, 24, 8);
            numMainBuf  = AddNum(right, "缓冲(ms)", 2, 500, 9, 1);

            chkMainStrict = AddCheck(right, "严格格式（不回退至32容器）");
            numMainAlignN = AddNum(right, "最小周期对齐倍数 N（0=自动）", 0, 16, 0, 1);

            // Aux group
            AddHeader(right, "副输出");
            chkAuxExclusive = AddCheck(right, "独占模式");
            chkAuxEvent = AddCheck(right, "事件驱动");
            numAuxRate = AddNum(right, "采样率(Hz)", 8000, 384000, 44100, 100);
            numAuxBits = AddNum(right, "位深(bit)", 16, 32, 16, 8);
            numAuxBuf  = AddNum(right, "缓冲(ms)", 2, 500, 120, 1);

            chkAuxStrict = AddCheck(right, "严格格式（不回退至32容器）");
            numAuxAlignN = AddNum(right, "最小周期对齐倍数 N（0=自动）", 0, 16, 0, 1);

            // Buttons
            btnApply = new Button { Text = "应用并重启音频", AutoSize = true };
            btnApply.Click += (_, __) => { SaveAndApply(); };
            right.Controls.Add(new Label()); // spacer
            right.Controls.Add(btnApply);

            btnSaveExit = new Button { Text = "保存并退出", AutoSize = true };
            btnSaveExit.Click += (_, __) => { SaveAndClose(); };
            right.Controls.Add(new Label());
            right.Controls.Add(btnSaveExit);
        }

        Label AddKV(TableLayoutPanel host, string key)
        {
            var row = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var k = new Label { Text = key, AutoSize = true, Font = new Font(FontFamily.GenericSansSerif, 9, FontStyle.Bold) };
            var v = new Label { Text = "-", AutoSize = true };
            row.Controls.Add(k, 0, 0);
            row.Controls.Add(v, 1, 0);
            host.Controls.Add(row);
            return v;
        }

        void AddHeader(TableLayoutPanel host, string text)
        {
            var lbl = new Label { Text = text, AutoSize = true, Font = new Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold), Margin = new Padding(0, 10, 0, 6) };
            host.Controls.Add(lbl);
            host.Controls.Add(new Label()); // occupy second column
        }

        CheckBox AddCheck(TableLayoutPanel host, string text)
        {
            var cb = new CheckBox { Text = text, AutoSize = true };
            host.Controls.Add(new Label()); // label column empty
            host.Controls.Add(cb);
            return cb;
        }

        NumericUpDown AddNum(TableLayoutPanel host, string label, decimal min, decimal max, decimal val, decimal inc)
        {
            var l = new Label { Text = label, AutoSize = true };
            var n = new NumericUpDown { Minimum = min, Maximum = max, Value = val, Increment = inc, DecimalPlaces = 0, ThousandsSeparator = true };
            host.Controls.Add(l);
            host.Controls.Add(n);
            return n;
        }

        public void RenderStatus()
        {
            var s = _getStatus();
            lblMainMode.Text = s.MainMode;
            lblMainFmt.Text  = s.MainFmt;
            lblMainMap.Text  = s.MainBitDepthMap;
            lblMainBuf.Text  = s.MainBuf;
            lblMainRes.Text  = s.MainResample;

            lblAuxMode.Text = s.AuxMode;
            lblAuxFmt.Text  = s.AuxFmt;
            lblAuxMap.Text  = s.AuxBitDepthMap;
            lblAuxBuf.Text  = s.AuxBuf;
            lblAuxRes.Text  = s.AuxResample;
        }

        void LoadConfig()
        {
            var c = _getConfig();

            chkMainExclusive.Checked = c.MainExclusive;
            chkMainEvent.Checked = c.MainEventDriven;
            numMainRate.Value = c.MainRate;
            numMainBits.Value = c.MainBits;
            numMainBuf.Value  = c.MainBufMs;

            chkMainStrict.Checked = c.MainStrictFormat;
            numMainAlignN.Value = c.MainAlignMultiple;

            chkAuxExclusive.Checked = c.AuxExclusive;
            chkAuxEvent.Checked = c.AuxEventDriven;
            numAuxRate.Value = c.AuxRate;
            numAuxBits.Value = c.AuxBits;
            numAuxBuf.Value  = c.AuxBufMs;

            chkAuxStrict.Checked = c.AuxStrictFormat;
            numAuxAlignN.Value = c.AuxAlignMultiple;
        }

        void SaveAndApply()
        {
            var c = _getConfig();
            c.MainExclusive = chkMainExclusive.Checked;
            c.MainEventDriven = chkMainEvent.Checked;
            c.MainRate = (int)numMainRate.Value;
            c.MainBits = (int)numMainBits.Value;
            c.MainBufMs = (int)numMainBuf.Value;
            c.MainStrictFormat = chkMainStrict.Checked;
            c.MainAlignMultiple = (int)numMainAlignN.Value;

            c.AuxExclusive = chkAuxExclusive.Checked;
            c.AuxEventDriven = chkAuxEvent.Checked;
            c.AuxRate = (int)numAuxRate.Value;
            c.AuxBits = (int)numAuxBits.Value;
            c.AuxBufMs = (int)numAuxBuf.Value;
            c.AuxStrictFormat = chkAuxStrict.Checked;
            c.AuxAlignMultiple = (int)numAuxAlignN.Value;

            _saveConfig(c);
            _applyAndRestart();
            RenderStatus();
        }

        void SaveAndClose()
        {
            var c = _getConfig();
            c.MainExclusive = chkMainExclusive.Checked;
            c.MainEventDriven = chkMainEvent.Checked;
            c.MainRate = (int)numMainRate.Value;
            c.MainBits = (int)numMainBits.Value;
            c.MainBufMs = (int)numMainBuf.Value;
            c.MainStrictFormat = chkMainStrict.Checked;
            c.MainAlignMultiple = (int)numMainAlignN.Value;

            c.AuxExclusive = chkAuxExclusive.Checked;
            c.AuxEventDriven = chkAuxEvent.Checked;
            c.AuxRate = (int)numAuxRate.Value;
            c.AuxBits = (int)numAuxBits.Value;
            c.AuxBufMs = (int)numAuxBuf.Value;
            c.AuxStrictFormat = chkAuxStrict.Checked;
            c.AuxAlignMultiple = (int)numAuxAlignN.Value;

            _saveConfig(c);
            this.Close();
        }
    }
}
