public SettingsForm(AppSettings cur, Func<StatusSnapshot> statusProvider)
{
    _statusProvider = statusProvider ?? (() => new StatusSnapshot { Running = false });

    Text = "MirrorAudio 设置";
    StartPosition = FormStartPosition.CenterScreen;
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = SystemFonts.MessageBoxFont;
    MinimumSize = new Size(980, 620);
    Size = new Size(1100, 680);

    // —— 左右各一半 —— //
    var split = new SplitContainer
    {
        Dock = DockStyle.Fill,
        Orientation = Orientation.Vertical,
        FixedPanel = FixedPanel.None,
        SplitterWidth = 6
    };
    Controls.Add(split);

    // 在 Shown 与 Resize 时保持 50/50
    EventHandler do50 = (s, e) =>
    {
        if (split.Width > 0) split.SplitterDistance = split.Width / 2;
    };
    Shown += do50; Resize += do50;

    // 左：状态（滚动 + 紧凑表格）
    var left = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
    var grpS = new GroupBox { Text = "当前状态（打开查看，关闭即释放内存）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
    var tblS = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
    tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
    tblS.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));

    AddRow(tblS, "运行状态", lblRun);
    AddRow(tblS, "输入", lblInput);
    AddRow(tblS, "主通道", lblMain);
    AddRow(tblS, "主格式", lblMainFmt);
    AddRow(tblS, "主缓冲", lblMainBuf);
    AddRow(tblS, "主周期", lblMainPer);
    AddRow(tblS, "副通道", lblAux);
    AddRow(tblS, "副格式", lblAuxFmt);
    AddRow(tblS, "副缓冲", lblAuxBuf);
    AddRow(tblS, "副周期", lblAuxPer);

    var pBtns = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 6, 0, 4) };
    btnRefresh.Text = "刷新状态";
    btnCopy.Text = "复制状态";
    btnRefresh.Click += (s, e) => RenderStatus();
    btnCopy.Click += (s, e) => { Clipboard.SetText(BuildStatusText()); MessageBox.Show("状态已复制到剪贴板。", "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Information); };
    pBtns.Controls.Add(btnRefresh);
    pBtns.Controls.Add(btnCopy);

    grpS.Controls.Add(tblS);
    grpS.Controls.Add(pBtns);
    left.Controls.Add(grpS);
    split.Panel1.Controls.Add(left);

    // 右：设置（顺序：输入 → 主通道 → 副通道；滚动）
    var right = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };

    // ① 输入
    var gInput = new GroupBox { Text = "输入（通道1：录音或环回）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
    var tInput = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
    tInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
    tInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
    cmbInput.DropDownStyle = ComboBoxStyle.DropDownList;
    AddRow(tInput, "输入设备", cmbInput);
    gInput.Controls.Add(tInput);
    right.Controls.Add(gInput);

    // ② 主通道
    var gMain = new GroupBox { Text = "主通道（低延迟，高音质）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
    var tMain = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
    tMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
    tMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
    cmbMain.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbShareMain.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbSyncMain.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbShareMain.Items.AddRange(new object[] { "自动（优先独占）", "强制独占", "强制共享" });
    cmbSyncMain.Items.AddRange(new object[] { "自动（事件优先）", "强制事件", "强制轮询" });
    numRateMain.Maximum = 384000; numRateMain.Minimum = 44100;  numRateMain.Increment = 1000;
    numBitsMain.Maximum = 32;     numBitsMain.Minimum = 16;     numBitsMain.Increment = 8;
    numBufMain.Maximum  = 200;    numBufMain.Minimum  = 4;

    AddRow(tMain, "主输出设备", cmbMain);
    AddRow(tMain, "模式", cmbShareMain);
    AddRow(tMain, "同步方式", cmbSyncMain);
    AddRow(tMain, "采样率 (Hz，仅独占)", numRateMain);
    AddRow(tMain, "位深 (bit，仅独占)",  numBitsMain);
    AddRow(tMain, "缓冲 (ms)",            numBufMain);
    gMain.Controls.Add(tMain);
    right.Controls.Add(gMain);

    // ③ 副通道
    var gAux = new GroupBox { Text = "副通道（推流，稳定优先）", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
    var tAux = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
    tAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
    tAux.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
    cmbAux.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbShareAux.DropDownStyle = ComboBoxStyle.DropDownList;
    cmbSyncAux.DropDownStyle  = ComboBoxStyle.DropDownList;
    cmbShareAux.Items.AddRange(new object[] { "自动（优先独占）", "强制独占", "强制共享" });
    cmbSyncAux .Items.AddRange(new object[] { "自动（事件优先）", "强制事件", "强制轮询" });
    numRateAux.Maximum = 384000;  numRateAux.Minimum = 44100;   numRateAux.Increment = 1000;
    numBitsAux.Maximum = 32;      numBitsAux.Minimum = 16;      numBitsAux.Increment = 8;
    numBufAux.Maximum  = 400;     numBufAux.Minimum  = 50;

    AddRow(tAux, "副输出设备", cmbAux);
    AddRow(tAux, "模式", cmbShareAux);
    AddRow(tAux, "同步方式", cmbSyncAux);
    AddRow(tAux, "采样率 (Hz，仅独占)", numRateAux);
    AddRow(tAux, "位深 (bit，仅独占)",  numBitsAux);
    AddRow(tAux, "缓冲 (ms)",            numBufAux);
    gAux.Controls.Add(tAux);
    right.Controls.Add(gAux);

    // ④ 其他
    var gOpt = new GroupBox { Text = "其他", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
    var pOpt = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, AutoSize = true };
    chkAutoStart.Text = "Windows 自启动";
    chkLogging.Text   = "启用日志（排障时开启）";
    pOpt.Controls.Add(chkAutoStart); pOpt.Controls.Add(chkLogging);
    gOpt.Controls.Add(pOpt);
    right.Controls.Add(gOpt);

    split.Panel2.Controls.Add(right);

    // 底部按钮
    var pnlButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Padding = new Padding(10), AutoSize = true };
    btnOk.Text = "保存"; btnCancel.Text = "取消";
    AcceptButton = btnOk; CancelButton = btnCancel;
    btnOk.DialogResult = DialogResult.OK; btnCancel.DialogResult = DialogResult.Cancel;
    btnOk.Click += (s, e) => SaveAndClose();
    pnlButtons.Controls.AddRange(new Control[] { btnOk, btnCancel });
    Controls.Add(pnlButtons);

    // 加载设备与配置 + 渲染状态
    LoadDevices();
    LoadConfig(cur);
    RenderStatus();
}
