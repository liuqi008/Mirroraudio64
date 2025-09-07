using System;
using System.Linq;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace MirrorAudio.AppContextApp
{
    public sealed class SettingsForm : Form
    {
        ComboBox cmbInput = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 360 };
        ComboBox cmbMain = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 360 };
        ComboBox cmbAux = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 360 };

        CheckBox chkMainExclusive = new CheckBox { Text = "主通道独占(Exclusive)" };
        CheckBox chkMainRaw = new CheckBox { Text = "主通道RAW直通（绕开APM）" };
        CheckBox chkMainForce = new CheckBox { Text = "主通道强制使用下方格式" };
        NumericUpDown numMainRate = new NumericUpDown { Minimum = 8000, Maximum = 384000, Increment = 10, Value = 192000, Width = 100 };
        NumericUpDown numMainBits = new NumericUpDown { Minimum = 16, Maximum = 32, Increment = 8, Value = 24, Width = 100 };
        NumericUpDown numMainBuf = new NumericUpDown { Minimum = 2, Maximum = 1000, Increment = 1, Value = 12, Width = 100 };

        CheckBox chkAuxExclusive = new CheckBox { Text = "副通道独占(Exclusive)" };
        CheckBox chkAuxRaw = new CheckBox { Text = "副通道RAW直通（绕开APM）" };
        CheckBox chkAuxForce = new CheckBox { Text = "副通道强制使用下方格式" };
        NumericUpDown numAuxRate = new NumericUpDown { Minimum = 8000, Maximum = 384000, Increment = 10, Value = 44100, Width = 100 };
        NumericUpDown numAuxBits = new NumericUpDown { Minimum = 16, Maximum = 32, Increment = 8, Value = 16, Width = 100 };
        NumericUpDown numAuxBuf = new NumericUpDown { Minimum = 2, Maximum = 1000, Increment = 1, Value = 120, Width = 100 };

        CheckBox chkAutoStart = new CheckBox { Text = "随系统自动启动" };
        Button btnSave = new Button { Text = "保存", Width = 120 };
        Button btnCancel = new Button { Text = "取消", Width = 120 };

        public SettingsForm()
        {
            Text = "MirrorAudio 设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            var gInput = new GroupBox { Text = "输入源（留空=系统环回）", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12) };
            gInput.Controls.Add(new FlowLayoutPanel { AutoSize = true, Controls = { new Label { Text = "输入设备：" }, cmbInput } });

            var gMain = new GroupBox { Text = "主通道（SPDIF）", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12) };
            gMain.Controls.Add(new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Controls = {
                new FlowLayoutPanel { AutoSize=true, Controls = { new Label{Text="输出设备："}, cmbMain } },
                chkMainExclusive, chkMainRaw, chkMainForce,
                new FlowLayoutPanel { AutoSize=true, Controls = { new Label{Text="采样率："}, numMainRate, new Label{Text="位深："}, numMainBits, new Label{Text="缓冲(ms)："}, numMainBuf } }
            }});

            var gAux = new GroupBox { Text = "副通道", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12) };
            gAux.Controls.Add(new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Controls = {
                new FlowLayoutPanel { AutoSize=true, Controls = { new Label{Text="输出设备："}, cmbAux } },
                chkAuxExclusive, chkAuxRaw, chkAuxForce,
                new FlowLayoutPanel { AutoSize=true, Controls = { new Label{Text="采样率："}, numAuxRate, new Label{Text="位深："}, numAuxBits, new Label{Text="缓冲(ms)："}, numAuxBuf } }
            }});

            var bottom = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Controls = { chkAutoStart, btnSave, btnCancel } };

            var root = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Controls = { gInput, gMain, gAux, bottom } };
            Controls.Add(root);

            btnSave.Click += (_, __) => { Save(); DialogResult = DialogResult.OK; Close(); };
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            Load += (_, __) => { LoadData(); };
        }

        private void LoadData()
        {
            var cfg = AppSettings.Load();
            using var mm = new MMDeviceEnumerator();

            cmbInput.Items.Clear();
            cmbInput.Items.Add(new DevItem(null, "<系统环回>"));
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                cmbInput.Items.Add(new DevItem(d.ID, $"{d.FriendlyName}"));

            cmbMain.Items.Clear();
            cmbAux.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                var item = new DevItem(d.ID, $"{d.FriendlyName}");
                cmbMain.Items.Add(item);
                cmbAux.Items.Add(new DevItem(d.ID, $"{d.FriendlyName}"));
            }

            SelectById(cmbInput, cfg.InputDeviceId);
            SelectById(cmbMain, cfg.MainDeviceId);
            SelectById(cmbAux, cfg.AuxDeviceId);

            chkMainExclusive.Checked = cfg.MainExclusive;
            chkMainRaw.Checked = cfg.MainRaw;
            chkMainForce.Checked = cfg.MainForceFormat;
            numMainRate.Value = cfg.MainSampleRate;
            numMainBits.Value = cfg.MainBits;
            numMainBuf.Value = cfg.MainBufferMs;

            chkAuxExclusive.Checked = cfg.AuxExclusive;
            chkAuxRaw.Checked = cfg.AuxRaw;
            chkAuxForce.Checked = cfg.AuxForceFormat;
            numAuxRate.Value = cfg.AuxSampleRate;
            numAuxBits.Value = cfg.AuxBits;
            numAuxBuf.Value = cfg.AuxBufferMs;

            chkAutoStart.Checked = cfg.AutoStart;
        }

        private static void SelectById(ComboBox cmb, string? id)
        {
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                if (((DevItem)cmb.Items[i]).Id == id)
                { cmb.SelectedIndex = i; return; }
            }
            if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
        }

        private void Save()
        {
            var cfg = AppSettings.Load();
            cfg.InputDeviceId = ((DevItem)cmbInput.SelectedItem).Id;
            cfg.MainDeviceId = ((DevItem)cmbMain.SelectedItem).Id;
            cfg.AuxDeviceId = ((DevItem)cmbAux.SelectedItem).Id;

            cfg.MainExclusive = chkMainExclusive.Checked;
            cfg.MainRaw = chkMainRaw.Checked;
            cfg.MainForceFormat = chkMainForce.Checked;
            cfg.MainSampleRate = (int)numMainRate.Value;
            cfg.MainBits = (int)numMainBits.Value;
            cfg.MainBufferMs = (int)numMainBuf.Value;

            cfg.AuxExclusive = chkAuxExclusive.Checked;
            cfg.AuxRaw = chkAuxRaw.Checked;
            cfg.AuxForceFormat = chkAuxForce.Checked;
            cfg.AuxSampleRate = (int)numAuxRate.Value;
            cfg.AuxBits = (int)numAuxBits.Value;
            cfg.AuxBufferMs = (int)numAuxBuf.Value;

            cfg.AutoStart = chkAutoStart.Checked;
            cfg.Save();
        }

        private sealed record DevItem(string? Id, string Text)
        {
            public override string ToString() => Text;
        }
    }
}