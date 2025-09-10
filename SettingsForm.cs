using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio
{
    sealed class SettingsForm : Form
    {
        // Left status
        readonly Label lblRun=new Label(),
                       lblInput=new Label(),
                       lblInputReq=new Label(),
                       lblMain=new Label(), lblAux=new Label(),
                       lblMainFmt=new Label(), lblAuxFmt=new Label(),
                       lblMainPass=new Label(), lblAuxPass=new Label(),
                       lblMainBuf=new Label(), lblAuxBuf=new Label(),
                       lblMainPer=new Label(), lblAuxPer=new Label(),
                       lblMainSRC2=new Label(), lblAuxSRC2=new Label(),
                       lblMainQ=new Label(), lblAuxQ=new Label();

        // Right settings
        readonly ComboBox cmbInput=new ComboBox(), cmbMain=new ComboBox(), cmbAux=new ComboBox();
        readonly ComboBox cmbShareMain=new ComboBox(), cmbSyncMain=new ComboBox(), cmbShareAux=new ComboBox(), cmbSyncAux=new ComboBox();
        readonly ComboBox cmbBufModeMain=new ComboBox(), cmbBufModeAux=new ComboBox();
        readonly NumericUpDown numRateMain=new NumericUpDown(), numBitsMain=new NumericUpDown(), numBufMain=new NumericUpDown();
        readonly NumericUpDown numRateAux=new NumericUpDown(), numBitsAux=new NumericUpDown(), numBufAux=new NumericUpDown();
        readonly ComboBox cmbResampMain=new ComboBox(), cmbResampAux=new ComboBox();
        readonly CheckBox chkMainForceInShared=new CheckBox(), chkAuxForceInShared=new CheckBox();

        // Input loopback
        readonly ComboBox cmbInStrategy=new ComboBox();
        readonly NumericUpDown numInRate=new NumericUpDown(), numInBits=new NumericUpDown();

        readonly CheckBox chkAutoStart=new CheckBox(), chkLogging=new CheckBox();
        readonly Button btnOk=new Button(), btnCancel=new Button(), btnRefresh=new Button(), btnReload=new Button();

        readonly Func<StatusSnapshot> _statusProvider;
        public AppSettings Result { get; private set; }
        sealed class DevItem { public string Id,Name; public override string ToString(){ return Name; } }

        public SettingsForm(AppSettings cur, Func<StatusSnapshot> statusProvider)
        {
            _statusProvider = statusProvider ?? (() => new StatusSnapshot { Running = false });
            Text = "MirrorAudio 设置"; StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi; Font = SystemFonts.MessageBoxFont;
            MinimumSize = new Size(1000, 680); Size = new Size(1160, 720);

            var split = new SplitContainer{ Dock=DockStyle.Fill, Orientation=Orientation.Vertical, SplitterWidth=6 };
            Controls.Add(split);
            EventHandler keepHalf=(s,e)=>{ if(split.Width>0) split.SplitterDistance=split.Width/2; };
            Shown+=keepHalf; Resize+=keepHalf;

            // Left panel (status)
            var left = new Panel{ Dock=DockStyle.Fill, AutoScroll=true, Padding=new Padding(10) };
            var grpS = new GroupBox{ Text="当前状态", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var tblS = new TableLayoutPanel{ ColumnCount=2, Dock=DockStyle.Top, AutoSize=true };
            tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
            tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
            AddRow(tblS, "运行状态", lblRun);
            AddRow(tblS, "输入", lblInput);
            AddRow(tblS, "环回请求/接受/混音", lblInputReq);
            AddRow(tblS, "主通道", lblMain);
            AddRow(tblS, "主格式", lblMainFmt);
            AddRow(tblS, "主直通/重采样", lblMainPass);
            AddRow(tblS, "主：程序内重采样/多次SRC", lblMainSRC2);
            AddRow(tblS, "主：程序内重采样质量", lblMainQ);
            AddRow(tblS, "主缓冲", lblMainBuf);
            AddRow(tblS, "主周期", lblMainPer);
            AddRow(tblS, "副通道", lblAux);
            AddRow(tblS, "副格式", lblAuxFmt);
            AddRow(tblS, "副直通/重采样", lblAuxPass);
            AddRow(tblS, "副：程序内重采样/多次SRC", lblAuxSRC2);
            AddRow(tblS, "副：程序内重采样质量", lblAuxQ);
            AddRow(tblS, "副缓冲", lblAuxBuf);
            AddRow(tblS, "副周期", lblAuxPer);
            grpS.Controls.Add(tblS);

            var pBtns = new FlowLayoutPanel{ FlowDirection=FlowDirection.LeftToRight, Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(0,6,0,4) };
            btnRefresh.Text="刷新状态"; btnRefresh.Click+=(s,e)=>{ RenderStatus(); UpdateResamplerControlsEnabled(); };
            pBtns.Controls.Add(btnRefresh); grpS.Controls.Add(pBtns);

            left.Controls.Add(grpS);
            split.Panel1.Controls.Add(left);

            // Right panel (settings)
            var right = new Panel{ Dock=DockStyle.Fill, AutoScroll=true, Padding=new Padding(10) };
            split.Panel2.Controls.Add(right);

            // Devices
            var gDev = new GroupBox{ Text="设备", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var tDev = new TableLayoutPanel{ ColumnCount=2, Dock=DockStyle.Top, AutoSize=true };
            tDev.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            tDev.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            cmbInput.DropDownStyle=ComboBoxStyle.DropDownList;
            cmbMain.DropDownStyle=ComboBoxStyle.DropDownList;
            cmbAux.DropDownStyle=ComboBoxStyle.DropDownList;
            AddRow(tDev, "通道1 输入设备", cmbInput);
            AddRow(tDev, "通道2 主输出设备", cmbMain);
            AddRow(tDev, "通道3 副输出设备", cmbAux);
            btnReload.Text="重新枚举设备"; btnReload.AutoSize=true; btnReload.Click+=(s,e)=>LoadDevices();
            tDev.RowStyles.Add(new RowStyle(SizeType.AutoSize)); tDev.Controls.Add(btnReload, 1, tDev.RowCount++);
            gDev.Controls.Add(tDev); right.Controls.Add(gDev);

            // Input strategy
            var gIn = new GroupBox{ Text="输入（环回）格式策略", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var tIn = new TableLayoutPanel{ ColumnCount=2, Dock=DockStyle.Top, AutoSize=true };
            tIn.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            tIn.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            cmbInStrategy.DropDownStyle=ComboBoxStyle.DropDownList;
            cmbInStrategy.Items.AddRange(new object[]{
                "System Mix（跟随系统混音）",
                "24-bit / 48 kHz",
                "24-bit / 96 kHz",
                "24-bit / 192 kHz",
                "32-float / 48 kHz",
                "32-float / 96 kHz",
                "32-float / 192 kHz",
                "自定义..."
            });
            numInRate.Minimum=8000; numInRate.Maximum=384000; numInRate.Increment=1000; numInRate.Width=140;
            numInBits.Minimum=16; numInBits.Maximum=32; numInBits.Increment=8; numInBits.Width=140;
            AddRow(tIn, "环回请求策略", cmbInStrategy);
            AddRow(tIn, "自定义采样率 (Hz)", numInRate);
            AddRow(tIn, "自定义位深 (16/24/32f)", numInBits);
            gIn.Controls.Add(tIn); right.Controls.Add(gIn);

            // Main output
            var gMain = new GroupBox{ Text="主输出（高音质，低延迟）", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var tMain = new TableLayoutPanel{ ColumnCount=2, Dock=DockStyle.Top, AutoSize=true };
            tMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            tMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            cmbShareMain.DropDownStyle=ComboBoxStyle.DropDownList; cmbShareMain.Items.AddRange(new object[]{"自动（优先独占）","强制独占","强制共享"});
            cmbSyncMain.DropDownStyle=ComboBoxStyle.DropDownList; cmbSyncMain.Items.AddRange(new object[]{"自动（事件优先）","强制事件","强制轮询"});
            numRateMain.Maximum=384000; numRateMain.Minimum=44100; numRateMain.Increment=1000;
            numBitsMain.Maximum=32; numBitsMain.Minimum=16; numBitsMain.Increment=8;
            numBufMain.Maximum=200; numBufMain.Minimum=4;
            AddRow(tMain, "模式", cmbShareMain);
            AddRow(tMain, "同步方式", cmbSyncMain);
            AddRow(tMain, "采样率 (Hz，仅独占)", numRateMain);
            AddRow(tMain, "位深 (bit，仅独占)", numBitsMain);
            AddRow(tMain, "缓冲 (ms)", numBufMain);
            cmbBufModeMain.DropDownStyle=ComboBoxStyle.DropDownList; cmbBufModeMain.Items.AddRange(new object[]{ "默认对齐", "最小对齐" });
            AddRow(tMain, "缓冲对齐模式", cmbBufModeMain);
            cmbResampMain.DropDownStyle=ComboBoxStyle.DropDownList; cmbResampMain.Items.AddRange(new object[]{ "60","50","40","30" });
            chkMainForceInShared.Text="共享模式下也程序内重采样";
            var pMainRow=new FlowLayoutPanel{ FlowDirection=FlowDirection.LeftToRight, AutoSize=true, Dock=DockStyle.Fill };
            pMainRow.Controls.Add(cmbResampMain); pMainRow.Controls.Add(chkMainForceInShared);
            AddRow(tMain, "重采样质量", pMainRow);
            gMain.Controls.Add(tMain); right.Controls.Add(gMain);

            // Aux output
            var gAux = new GroupBox{ Text="副输出（直播/采集卡）", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var tAux = new TableLayoutPanel{ ColumnCount=2, Dock=DockStyle.Top, AutoSize=true };
            tAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            tAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            cmbShareAux.DropDownStyle=ComboBoxStyle.DropDownList; cmbShareAux.Items.AddRange(new object[]{"自动（优先独占）","强制独占","强制共享"});
            cmbSyncAux.DropDownStyle=ComboBoxStyle.DropDownList;  cmbSyncAux.Items.AddRange(new object[]{"自动（事件优先）","强制事件","强制轮询"});
            numRateAux.Maximum=384000; numRateAux.Minimum=44100; numRateAux.Increment=1000;
            numBitsAux.Maximum=32; numBitsAux.Minimum=16; numBitsAux.Increment=8;
            numBufAux.Maximum=400; numBufAux.Minimum=50;
            AddRow(tAux, "模式", cmbShareAux);
            AddRow(tAux, "同步方式", cmbSyncAux);
            AddRow(tAux, "采样率 (Hz，仅独占)", numRateAux);
            AddRow(tAux, "位深 (bit，仅独占)", numBitsAux);
            AddRow(tAux, "缓冲 (ms)", numBufAux);
            cmbBufModeAux.DropDownStyle=ComboBoxStyle.DropDownList; cmbBufModeAux.Items.AddRange(new object[]{ "默认对齐", "最小对齐" });
            AddRow(tAux, "缓冲对齐模式", cmbBufModeAux);
            cmbResampAux.DropDownStyle=ComboBoxStyle.DropDownList; cmbResampAux.Items.AddRange(new object[]{ "60","50","40","30" });
            chkAuxForceInShared.Text="共享模式下也程序内重采样";
            var pAuxRow=new FlowLayoutPanel{ FlowDirection=FlowDirection.LeftToRight, AutoSize=true, Dock=DockStyle.Fill };
            pAuxRow.Controls.Add(cmbResampAux); pAuxRow.Controls.Add(chkAuxForceInShared);
            AddRow(tAux, "重采样质量", pAuxRow);
            gAux.Controls.Add(tAux); right.Controls.Add(gAux);

            // Options
            var gOpt=new GroupBox{ Text="其他", Dock=DockStyle.Top, AutoSize=true, Padding=new Padding(10) };
            var pOpt=new FlowLayoutPanel{ FlowDirection=FlowDirection.LeftToRight, Dock=DockStyle.Top, AutoSize=true };
            chkAutoStart.Text="Windows 自启动"; chkLogging.Text="启用日志（排障时开启）";
            pOpt.Controls.Add(chkAutoStart); pOpt.Controls.Add(chkLogging);
            gOpt.Controls.Add(pOpt); right.Controls.Add(gOpt);

            // Bottom buttons
            var pnlButtons=new FlowLayoutPanel{ FlowDirection=FlowDirection.RightToLeft, Dock=DockStyle.Bottom, Padding=new Padding(10), AutoSize=true };
            btnOk.Text="保存"; btnCancel.Text="取消"; AcceptButton=btnOk; CancelButton=btnCancel;
            btnOk.DialogResult=DialogResult.OK; btnCancel.DialogResult=DialogResult.Cancel;
            btnOk.Click+=(s,e)=>SaveAndClose();
            pnlButtons.Controls.AddRange(new Control[]{ btnOk, btnCancel });
            Controls.Add(pnlButtons);

            // Init
            LoadDevices();
            LoadConfig(cur);
            RenderStatus();
            UpdateResamplerControlsEnabled();

            cmbInStrategy.SelectedIndexChanged+=(s,e)=>{ bool cust = (cmbInStrategy.SelectedIndex == (int)InputFormatStrategy.Custom); numInRate.Enabled=cust; numInBits.Enabled=cust; };
            cmbShareMain.SelectedIndexChanged+=(s,e)=>UpdateResamplerControlsEnabled();
            cmbShareAux.SelectedIndexChanged +=(s,e)=>UpdateResamplerControlsEnabled();
            chkMainForceInShared.CheckedChanged +=(s,e)=>UpdateResamplerControlsEnabled();
            chkAuxForceInShared.CheckedChanged  +=(s,e)=>UpdateResamplerControlsEnabled();
        }

        // Load config
        void LoadConfig(AppSettings cur)
        {
            Result = new AppSettings
            {
                InputDeviceId = cur.InputDeviceId, MainDeviceId = cur.MainDeviceId, AuxDeviceId = cur.AuxDeviceId,
                MainShare = cur.MainShare, MainSync = cur.MainSync, MainRate = cur.MainRate, MainBits = cur.MainBits, MainBufMs = cur.MainBufMs,
                AuxShare = cur.AuxShare, AuxSync = cur.AuxSync, AuxRate = cur.AuxRate, AuxBits = cur.AuxBits, AuxBufMs = cur.AuxBufMs,
                AutoStart = cur.AutoStart, EnableLogging = cur.EnableLogging,
                InputFormatStrategy = cur.InputFormatStrategy, InputCustomSampleRate = cur.InputCustomSampleRate, InputCustomBitDepth = cur.InputCustomBitDepth,
                MainBufMode = cur.MainBufMode, AuxBufMode = cur.AuxBufMode,
                MainResamplerQuality = cur.MainResamplerQuality, AuxResamplerQuality = cur.AuxResamplerQuality,
                MainForceInternalResamplerInShared = cur.MainForceInternalResamplerInShared, AuxForceInternalResamplerInShared = cur.AuxForceInternalResamplerInShared
            };

            SelectById(cmbInput, cur.InputDeviceId);
            SelectById(cmbMain,  cur.MainDeviceId);
            SelectById(cmbAux,   cur.AuxDeviceId);

            cmbInStrategy.SelectedIndex = (int)cur.InputFormatStrategy;
            numInRate.Value = Clamp(cur.InputCustomSampleRate,  (int)numInRate.Minimum, (int)numInRate.Maximum);
            numInBits.Value = Clamp(cur.InputCustomBitDepth,    (int)numInBits.Minimum, (int)numInBits.Maximum);

            numRateMain.Value = Clamp(cur.MainRate, (int)numRateMain.Minimum, (int)numRateMain.Maximum);
            numBitsMain.Value = Clamp(cur.MainBits, (int)numBitsMain.Minimum, (int)numBitsMain.Maximum);
            numBufMain.Value  = Clamp(cur.MainBufMs,(int)numBufMain.Minimum,  (int)numBufMain.Maximum);
            cmbShareMain.SelectedIndex = cur.MainShare==ShareModeOption.Auto?0:(cur.MainShare==ShareModeOption.Exclusive?1:2);
            cmbSyncMain .SelectedIndex = cur.MainSync ==SyncModeOption .Auto?0:(cur.MainSync ==SyncModeOption .Event?1:2);
            cmbBufModeMain.SelectedIndex = cur.MainBufMode==BufferAlignMode.MinAlign ? 1 : 0;
            cmbResampMain.SelectedItem = (cur.MainResamplerQuality==0? "50" : cur.MainResamplerQuality.ToString());
            chkMainForceInShared.Checked = cur.MainForceInternalResamplerInShared;

            numRateAux.Value = Clamp(cur.AuxRate, (int)numRateAux.Minimum, (int)numRateAux.Maximum);
            numBitsAux.Value = Clamp(cur.AuxBits, (int)numBitsAux.Minimum, (int)numBitsAux.Maximum);
            numBufAux.Value  = Clamp(cur.AuxBufMs,(int)numBufAux.Minimum,  (int)numBufAux.Maximum);
            cmbShareAux.SelectedIndex = cur.AuxShare==ShareModeOption.Auto?0:(cur.AuxShare==ShareModeOption.Exclusive?1:2);
            cmbSyncAux .SelectedIndex = cur.AuxSync ==SyncModeOption .Auto?0:(cur.AuxSync ==SyncModeOption .Event?1:2);
            cmbBufModeAux.SelectedIndex = cur.AuxBufMode==BufferAlignMode.MinAlign ? 1 : 0;
            cmbResampAux.SelectedItem = (cur.AuxResamplerQuality==0? "30" : cur.AuxResamplerQuality.ToString());
            chkAuxForceInShared.Checked = cur.AuxForceInternalResamplerInShared;

            chkAutoStart.Checked = cur.AutoStart;
            chkLogging.Checked   = cur.EnableLogging;
        }

        void RenderStatus()
        {
            StatusSnapshot s; try { s = _statusProvider(); } catch { s = new StatusSnapshot(); }
            lblRun.Text   = s.Running ? "运行中" : "停止";
            lblInput.Text = (s.InputDevice??"-") + " | " + (s.InputRole??"-") + " | 实得: " + (s.InputFormat??"-");
            lblInputReq.Text = "请求: " + (s.InputRequested??"-") + "  |  接受: " + (s.InputAccepted??"-") + "  |  混音: " + (s.InputMix??"-");

            lblMain.Text  = (s.MainDevice ??"-") + " | " + (s.MainMode ??"-") + " | " + (s.MainSync ??"-");
            lblAux.Text   = (s.AuxDevice  ??"-") + " | " + (s.AuxMode  ??"-") + " | " + (s.AuxSync  ??"-");
            lblMainFmt.Text = s.MainFormat ?? "-";
            lblAuxFmt.Text  = s.AuxFormat  ?? "-";

            lblMainPass.Text = "直通=" + (s.MainNoSRC ? "是" : "否") + " | 重采样=" + (s.MainResampling ? "是" : "否");
            lblAuxPass.Text  = "直通=" + (s.AuxNoSRC ? "是" : "否") + " | 重采样=" + (s.AuxResampling ? "是" : "否");

            lblMainSRC2.Text = "程序内重采样=" + (s.MainInternalResampler ? "是" : "否") + "  |  多次重采样 SRC=" + (s.MainMultiSRC ? "是" : "否");
            lblAuxSRC2.Text  = "程序内重采样=" + (s.AuxInternalResampler  ? "是" : "否") + "  |  多次重采样 SRC=" + (s.AuxMultiSRC  ? "是" : "否");

            lblMainQ.Text = (s.MainInternalResamplerQuality > 0) ? s.MainInternalResamplerQuality.ToString() : "未生效";
            lblAuxQ.Text  = (s.AuxInternalResamplerQuality  > 0) ? s.AuxInternalResamplerQuality .ToString() : "未生效";
        }

        void UpdateResamplerControlsEnabled()
        {
            var s = _statusProvider != null ? _statusProvider() : null;

            bool mainShared = cmbShareMain.SelectedIndex==2; // Shared
            bool mainForce  = chkMainForceInShared.Checked;
            bool mainEnable = !mainShared || (mainShared && mainForce);
            if (s != null && s.Running && !s.MainInternalResampler) mainEnable = false; // 运行态未生效则灰掉
            cmbResampMain.Enabled = mainEnable;
            chkMainForceInShared.Enabled = mainShared; // 仅共享时有意义

            bool auxShared = cmbShareAux.SelectedIndex==2;
            bool auxForce  = chkAuxForceInShared.Checked;
            bool auxEnable = !auxShared || (auxShared && auxForce);
            if (s != null && s.Running && !s.AuxInternalResampler) auxEnable = false;
            cmbResampAux.Enabled = auxEnable;
            chkAuxForceInShared.Enabled = auxShared;
        }

        void SaveAndClose()
        {
            var inSel   = cmbInput.SelectedItem as DevItem;
            var mainSel = cmbMain .SelectedItem as DevItem;
            var auxSel  = cmbAux  .SelectedItem as DevItem;
            if (mainSel == null || auxSel == null)
            {
                MessageBox.Show("请至少选择主/副输出设备。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }

            Result = new AppSettings();
            Result.InputDeviceId = inSel != null ? inSel.Id : null;
            Result.MainDeviceId  = mainSel.Id;
            Result.AuxDeviceId   = auxSel.Id;

            Result.MainShare = cmbShareMain.SelectedIndex==1? ShareModeOption.Exclusive : (cmbShareMain.SelectedIndex==2? ShareModeOption.Shared : ShareModeOption.Auto);
            Result.MainSync  = cmbSyncMain .SelectedIndex==1? SyncModeOption .Event     : (cmbSyncMain .SelectedIndex==2? SyncModeOption .Polling : SyncModeOption .Auto);
            Result.AuxShare  = cmbShareAux .SelectedIndex==1? ShareModeOption.Exclusive : (cmbShareAux .SelectedIndex==2? ShareModeOption.Shared : ShareModeOption.Auto);
            Result.AuxSync   = cmbSyncAux  .SelectedIndex==1? SyncModeOption .Event     : (cmbSyncAux  .SelectedIndex==2? SyncModeOption .Polling : SyncModeOption .Auto);

            Result.MainRate = (int)numRateMain.Value;
            Result.MainBits = (int)numBitsMain.Value;
            Result.MainBufMs = (int)numBufMain.Value;
            Result.AuxRate = (int)numRateAux.Value;
            Result.AuxBits = (int)numBitsAux.Value;
            Result.AuxBufMs = (int)numBufAux.Value;

            Result.MainBufMode = cmbBufModeMain.SelectedIndex==1 ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign;
            Result.AuxBufMode  = cmbBufModeAux .SelectedIndex==1 ? BufferAlignMode.MinAlign : BufferAlignMode.DefaultAlign;

            Result.InputFormatStrategy = (InputFormatStrategy)cmbInStrategy.SelectedIndex;
            Result.InputCustomSampleRate = (int)numInRate.Value;
            Result.InputCustomBitDepth  = (int)numInBits.Value;

            int q;
            Result.MainResamplerQuality = int.TryParse((string)(cmbResampMain.SelectedItem ?? "50"), out q) ? q : 50;
            Result.AuxResamplerQuality  = int.TryParse((string)(cmbResampAux .SelectedItem ?? "30"), out q) ? q : 30;
            Result.MainForceInternalResamplerInShared = chkMainForceInShared.Checked;
            Result.AuxForceInternalResamplerInShared  = chkAuxForceInShared .Checked;

            Result.AutoStart = chkAutoStart.Checked;
            Result.EnableLogging = chkLogging.Checked;
        }

        // Helpers
        static void AddRow(TableLayoutPanel t, string label, Control c)
        {
            var l = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 6) };
            c.Dock = DockStyle.Fill;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(l, 0, t.RowCount);
            t.Controls.Add(c, 1, t.RowCount);
            t.RowCount++;
        }
        void LoadDevices()
        {
            var mm = new MMDeviceEnumerator();
            cmbInput.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="录音: " + d.FriendlyName });
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="环回: " + d.FriendlyName });
            cmbMain.Items.Clear(); cmbAux.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                cmbMain.Items.Add(new DevItem{ Id=d.ID, Name=d.FriendlyName });
                cmbAux .Items.Add(new DevItem{ Id=d.ID, Name=d.FriendlyName });
            }
            if (cmbInput.Items.Count>0 && cmbInput.SelectedIndex<0) cmbInput.SelectedIndex = 0;
            if (cmbMain .Items.Count>0 && cmbMain .SelectedIndex<0) cmbMain .SelectedIndex = 0;
            if (cmbAux  .Items.Count>0 && cmbAux  .SelectedIndex<0) cmbAux  .SelectedIndex = 0;
        }
        void SelectById(ComboBox cmb, string id)
        {
            if (string.IsNullOrEmpty(id) || cmb.Items.Count == 0) { if (cmb.Items.Count > 0) cmb.SelectedIndex = 0; return; }
            for (int i=0;i<cmb.Items.Count;i++){ var it=cmb.Items[i] as DevItem; if (it!=null && it.Id==id){ cmb.SelectedIndex=i; return; } }
            cmb.SelectedIndex = 0;
        }
        static int Clamp(int v, int lo, int hi){ return v<lo?lo:(v>hi?hi:v); }
    }
}
