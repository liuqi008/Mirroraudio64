using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    public class SettingsForm : Form
    {
        ComboBox cmbInput = new ComboBox();
        ComboBox cmbMain  = new ComboBox();
        ComboBox cmbAux   = new ComboBox();

        ComboBox cmbShare = new ComboBox(); // 主通道：共享/独占/自动
        ComboBox cmbSync  = new ComboBox(); // 主通道：事件/轮询/自动

        NumericUpDown numRate    = new NumericUpDown();
        NumericUpDown numBits    = new NumericUpDown();
        NumericUpDown numBufMain = new NumericUpDown();
        NumericUpDown numBufAux  = new NumericUpDown();

        CheckBox chkAutoStart = new CheckBox();
        CheckBox chkLogging   = new CheckBox();

        Button btnOk = new Button();
        Button btnCancel = new Button();

        class DevItem { public string Id; public string Name; public override string ToString() => Name; }

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings cur)
        {
            Text = "MirrorAudio 设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = SystemFonts.MessageBoxFont;
            ClientSize = new Size(720, 460);

            var title = new Label { Text = "选择设备并设定主通道策略（共享/独占、事件/轮询）。", AutoSize = true, ForeColor = SystemColors.GrayText, Padding = new Padding(12, 8, 12, 8) };

            var grid = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12, 0, 12, 0) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));

            foreach (var cb in new[] { cmbInput, cmbMain, cmbAux, cmbShare, cmbSync }) cb.DropDownStyle = ComboBoxStyle.DropDownList;

            AddRow(grid, "通道1 输入（录音/环回）", cmbInput);
            AddRow(grid, "通道2 主通道（低延迟）",   cmbMain);
            AddRow(grid, "通道3 副通道（省资源）",   cmbAux);

            cmbShare.Items.AddRange(new object[] { "自动（优先独占）", "强制独占", "强制共享" });
            cmbSync.Items.AddRange(new object[]  { "自动（事件优先）", "强制事件", "强制轮询" });
            AddRow(grid, "主通道模式", cmbShare);
            AddRow(grid, "主通道同步方式", cmbSync);

            numRate.Maximum = 384000; numRate.Minimum = 44100; numRate.Increment = 1000;
            numBits.Maximum = 32;     numBits.Minimum = 16;    numBits.Increment = 8;
            numBufMain.Maximum = 200; numBufMain.Minimum = 4;
            numBufAux.Maximum  = 400; numBufAux.Minimum  = 50;

            AddRow(grid, "主通道采样率 (Hz)",   numRate);
            AddRow(grid, "主通道路位深 (bit)",  numBits);
            AddRow(grid, "主通道缓冲 (ms)",     numBufMain);
            AddRow(grid, "副通道缓冲 (ms)",     numBufAux);

            var opts = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(18,6,12,6) };
            chkAutoStart.Text = "Windows 自启动";
            chkLogging.Text   = "启用日志（排障用）";
            opts.Controls.Add(chkAutoStart);
            opts.Controls.Add(chkLogging);

            var pnlButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Padding = new Padding(12), AutoSize = true };
            btnOk.Text = "保存"; btnCancel.Text = "取消";
            AcceptButton = btnOk; CancelButton = btnCancel;
            btnOk.DialogResult = DialogResult.OK; btnCancel.DialogResult = DialogResult.Cancel;
            btnOk.Click += (s,e)=> SaveAndClose();
            pnlButtons.Controls.AddRange(new Control[]{ btnOk, btnCancel });

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            root.Controls.Add(title, 0, 0);
            root.Controls.Add(grid,  0, 1);
            root.Controls.Add(opts,  0, 2);
            root.Controls.Add(pnlButtons, 0, 3);
            Controls.Add(root);

            LoadDevices();

            // 载入现有配置
            Result = new AppSettings {
                InputDeviceId = cur.InputDeviceId, MainDeviceId = cur.MainDeviceId, AuxDeviceId = cur.AuxDeviceId,
                MainShare = cur.MainShare, MainSync = cur.MainSync,
                MainRate = cur.MainRate, MainBits = cur.MainBits,
                MainBufMs = cur.MainBufMs, AuxBufMs = cur.AuxBufMs,
                AutoStart = cur.AutoStart, EnableLogging = cur.EnableLogging
            };

            SelectById(cmbInput, cur.InputDeviceId);
            SelectById(cmbMain,  cur.MainDeviceId);
            SelectById(cmbAux,   cur.AuxDeviceId);

            numRate.Value    = Clamp(cur.MainRate, (int)numRate.Minimum, (int)numRate.Maximum);
            numBits.Value    = Clamp(cur.MainBits, (int)numBits.Minimum, (int)numBits.Maximum);
            numBufMain.Value = Clamp(cur.MainBufMs,(int)numBufMain.Minimum,(int)numBufMain.Maximum);
            numBufAux.Value  = Clamp(cur.AuxBufMs, (int)numBufAux.Minimum, (int)numBufAux.Maximum);

            cmbShare.SelectedIndex = cur.MainShare == ShareModeOption.Auto ? 0 : (cur.MainShare == ShareModeOption.Exclusive ? 1 : 2);
            cmbSync.SelectedIndex  = cur.MainSync  == SyncModeOption.Auto  ? 0 : (cur.MainSync  == SyncModeOption.Event     ? 1 : 2);

            chkAutoStart.Checked = cur.AutoStart;
            chkLogging.Checked   = cur.EnableLogging;
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

        static int Clamp(int v, int lo, int hi) => v < lo ? lo : (v > hi ? hi : v);

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
                if ((cmb.Items[i] as DevItem)?.Id == id) { cmb.SelectedIndex = i; return; }
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

            ShareModeOption share = cmbShare.SelectedIndex == 1 ? ShareModeOption.Exclusive : (cmbShare.SelectedIndex == 2 ? ShareModeOption.Shared : ShareModeOption.Auto);
            SyncModeOption  sync  = cmbSync.SelectedIndex  == 1 ? SyncModeOption.Event      : (cmbSync.SelectedIndex  == 2 ? SyncModeOption.Polling : SyncModeOption.Auto);

            Result = new AppSettings {
                InputDeviceId = inSel != null ? inSel.Id : null,
                MainDeviceId  = mainSel.Id,
                AuxDeviceId   = auxSel.Id,
                MainShare     = share,
                MainSync      = sync,
                MainRate      = (int)numRate.Value,
                MainBits      = (int)numBits.Value,
                MainBufMs     = (int)numBufMain.Value,
                AuxBufMs      = (int)numBufAux.Value,
                AutoStart     = chkAutoStart.Checked,
                EnableLogging = chkLogging.Checked
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
