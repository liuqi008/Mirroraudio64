
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MirrorAudio
{
    public class SettingsForm : Form
    {
        // 基础控件
        readonly NumericUpDown numBufMain = new NumericUpDown();
        readonly NumericUpDown numBufAux  = new NumericUpDown();
        readonly ComboBox cmbShareMain = new ComboBox();
        readonly ComboBox cmbShareAux  = new ComboBox();
        readonly ComboBox cmbSyncMain  = new ComboBox();
        readonly ComboBox cmbSyncAux   = new ComboBox();

        // 本次新增/整合
        readonly ComboBox cmbBufModeMain = new ComboBox();
        readonly ComboBox cmbBufModeAux  = new ComboBox();
        readonly CheckBox chkAutoFallback = new CheckBox();

        readonly Button btnOK = new Button();
        readonly Button btnCancel = new Button();

        readonly AppSettings _cur;
        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings cur)
        {
            _cur = cur ?? new AppSettings();
            Text = "设置";
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            var table = new TableLayoutPanel{
                ColumnCount = 2, RowCount = 10,
                Dock = DockStyle.Fill, AutoSize = true,
                Padding = new Padding(0), CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // 行 1：主缓冲 (ms)
            numBufMain.Minimum = 3; numBufMain.Maximum = 500; numBufMain.Increment = 1;
            AddRow(table, "主缓冲 (ms)", numBufMain);

            // 行 2：主共享/独占
            cmbShareMain.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbShareMain.Items.AddRange(new object[]{ "独占", "共享" });
            AddRow(table, "主通道共享模式", cmbShareMain);

            // 行 3：主同步方式
            cmbSyncMain.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSyncMain.Items.AddRange(new object[]{ "轮询", "事件", "自动" });
            AddRow(table, "主通道同步方式", cmbSyncMain);

            // 行 4：主缓冲对齐模式
            cmbBufModeMain.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbBufModeMain.Items.AddRange(new object[]{ "默认对齐", "最小对齐" });
            AddRow(table, "主缓冲对齐模式", cmbBufModeMain);

            // 行 5：副缓冲 (ms)
            numBufAux.Minimum = 20; numBufAux.Maximum = 1000; numBufAux.Increment = 1;
            AddRow(table, "副缓冲 (ms)", numBufAux);

            // 行 6：副共享/独占
            cmbShareAux.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbShareAux.Items.AddRange(new object[]{ "独占", "共享" });
            AddRow(table, "副通道共享模式", cmbShareAux);

            // 行 7：副同步方式
            cmbSyncAux.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSyncAux.Items.AddRange(new object[]{ "轮询", "事件", "自动" });
            AddRow(table, "副通道同步方式", cmbSyncAux);

            // 行 8：副缓冲对齐模式
            cmbBufModeAux.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbBufModeAux.Items.AddRange(new object[]{ "默认对齐", "最小对齐" });
            AddRow(table, "副缓冲对齐模式", cmbBufModeAux);

            // 行 9：稳态保护
            chkAutoFallback.Text = "主路自动稳态回退（5秒内欠供≥2次→轮询）";
            chkAutoFallback.AutoSize = true;
            AddRow(table, "稳态保护", chkAutoFallback);

            // 行 10：按钮
            var btnPanel = new FlowLayoutPanel{ AutoSize=true, FlowDirection=FlowDirection.RightToLeft, Dock=DockStyle.Fill };
            btnOK.Text = "确定"; btnCancel.Text = "取消";
            btnOK.Click += OnOK; btnCancel.Click += (s,e)=>{ DialogResult=DialogResult.Cancel; Close(); };
            btnPanel.Controls.AddRange(new Control[]{ btnOK, btnCancel });
            AddRow(table, "", btnPanel);

            Controls.Add(table);

            // 载入当前配置
            LoadFrom(_cur);
        }

        void AddRow(TableLayoutPanel t, string label, Control ctrl)
        {
            var lab = new Label{ Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0,6,12,6) };
            ctrl.Margin = new Padding(0,3,0,3);
            int r = t.RowCount++;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(lab, 0, r);
            t.Controls.Add(ctrl, 1, r);
        }

        void LoadFrom(AppSettings cur)
        {
            numBufMain.Value = Math.Max(3, Math.Min(500, cur.MainBufMs));
            numBufAux.Value  = Math.Max(20, Math.Min(1000, cur.AuxBufMs));

            cmbShareMain.SelectedIndex = (cur.MainShare==0 ? 0 : 1);
            cmbShareAux.SelectedIndex  = (cur.AuxShare==0 ? 0 : 1);

            cmbSyncMain.SelectedIndex = Math.Max(0, Math.Min(2, cur.MainSync));
            cmbSyncAux.SelectedIndex  = Math.Max(0, Math.Min(2, cur.AuxSync));

            cmbBufModeMain.SelectedIndex = (cur.MainBufMode==BufferAlignMode.MinAlign ? 1 : 0);
            cmbBufModeAux.SelectedIndex  = (cur.AuxBufMode ==BufferAlignMode.MinAlign ? 1 : 0);

            chkAutoFallback.Checked = cur.AutoSyncFallback;
        }

        void OnOK(object sender, EventArgs e)
        {
            Result = new AppSettings{
                MainBufMs = (int)numBufMain.Value,
                AuxBufMs  = (int)numBufAux.Value,

                MainShare = cmbShareMain.SelectedIndex==0 ? 0 : 1,
                AuxShare  = cmbShareAux.SelectedIndex==0 ? 0 : 1,

                MainSync = cmbSyncMain.SelectedIndex,
                AuxSync  = cmbSyncAux.SelectedIndex,

                MainBufMode = (cmbBufModeMain.SelectedIndex==1 ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign),
                AuxBufMode  = (cmbBufModeAux .SelectedIndex==1 ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign),

                AutoSyncFallback = chkAutoFallback.Checked,

                // 保持你工程已有默认：
                MainBits=24, AuxBits=24,
                InputFormatStrategy= _cur.InputFormatStrategy,
                InputCustomSampleRate=_cur.InputCustomSampleRate,
                InputCustomBitDepth= _cur.InputCustomBitDepth,
                EnableLogging = _cur.EnableLogging,
                AutoStart = _cur.AutoStart
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
