using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MirrorAudio
{
    public partial class SettingsForm : Form
    {
        private readonly AppSettings _cfg;
        private readonly Func<StatusSnapshot> _getStatusSnapshot;

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings cfg, Func<StatusSnapshot> getStatusSnapshot)
        {
            _cfg = cfg;
            _getStatusSnapshot = getStatusSnapshot;
            InitializeComponent();
            PopulateSettings();
        }

        private void PopulateSettings()
        {
            // 设置界面填充默认值
            txtMainDevice.Text = _cfg.MainDeviceId;
            txtAuxDevice.Text = _cfg.AuxDeviceId;
            txtInputDevice.Text = _cfg.InputDeviceId;

            cmbMainSync.SelectedIndex = (int)_cfg.MainSync;
            cmbAuxSync.SelectedIndex = (int)_cfg.AuxSync;
            cmbMainShare.SelectedIndex = (int)_cfg.MainShare;
            cmbAuxShare.SelectedIndex = (int)_cfg.AuxShare;
            
            numMainRate.Value = _cfg.MainRate;
            numMainBits.Value = _cfg.MainBits;
            numMainBufMs.Value = _cfg.MainBufMs;

            numAuxRate.Value = _cfg.AuxRate;
            numAuxBits.Value = _cfg.AuxBits;
            numAuxBufMs.Value = _cfg.AuxBufMs;

            chkEnableLogging.Checked = _cfg.EnableLogging;

            // 状态区域填充
            var status = _getStatusSnapshot();
            lblStatus.Text = $"输入设备: {status.InputDevice}\n" +
                             $"输出设备: {status.MainDevice}\n" +
                             $"主通道状态: {status.MainMode}  ({status.MainSync})\n" +
                             $"副通道状态: {status.AuxMode}  ({status.AuxSync})\n" +
                             $"缓冲大小: 主({status.MainBufferMs}ms), 副({status.AuxBufferMs}ms)";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 保存设置
            _cfg.MainDeviceId = txtMainDevice.Text;
            _cfg.AuxDeviceId = txtAuxDevice.Text;
            _cfg.InputDeviceId = txtInputDevice.Text;
            _cfg.MainSync = (SyncModeOption)cmbMainSync.SelectedIndex;
            _cfg.AuxSync = (SyncModeOption)cmbAuxSync.SelectedIndex;
            _cfg.MainShare = (ShareModeOption)cmbMainShare.SelectedIndex;
            _cfg.AuxShare = (ShareModeOption)cmbAuxShare.SelectedIndex;
            
            _cfg.MainRate = (int)numMainRate.Value;
            _cfg.MainBits = (int)numMainBits.Value;
            _cfg.MainBufMs = (int)numMainBufMs.Value;

            _cfg.AuxRate = (int)numAuxRate.Value;
            _cfg.AuxBits = (int)numAuxBits.Value;
            _cfg.AuxBufMs = (int)numAuxBufMs.Value;

            _cfg.EnableLogging = chkEnableLogging.Checked;

            Result = _cfg;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // 关闭设置窗口
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
