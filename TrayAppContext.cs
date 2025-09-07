using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace MirrorAudio.AppContextApp
{
    public sealed class TrayAppContext : ApplicationContext
    {
        private readonly NotifyIcon _tray;
        private readonly ContextMenuStrip _menu;
        private readonly ToolStripMenuItem _startItem;
        private readonly ToolStripMenuItem _stopItem;
        private readonly ToolStripMenuItem _settingsItem;
        private readonly ToolStripMenuItem _exitItem;

        private AudioEngine? _engine;

        public TrayAppContext()
        {
            _menu = new ContextMenuStrip();
            _startItem = new ToolStripMenuItem("启动(&S)", null, (_, __) => Start());
            _stopItem = new ToolStripMenuItem("停止(&T)", null, (_, __) => Stop()) { Enabled = false };
            _settingsItem = new ToolStripMenuItem("设置(&E)...", null, (_, __) => ShowSettings());
            _exitItem = new ToolStripMenuItem("退出(&X)", null, (_, __) => ExitThread());

            _menu.Items.AddRange(new ToolStripItem[] { _startItem, _stopItem, new ToolStripSeparator(), _settingsItem, new ToolStripSeparator(), _exitItem });

            _tray = new NotifyIcon
            {
                Text = "MirrorAudio",
                Icon = new Icon(Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico")),
                Visible = true,
                ContextMenuStrip = _menu
            };
            _tray.DoubleClick += (_, __) => ShowSettings();

            if (AppSettings.Load().AutoStart)
            {
                // 延迟启动，避免首次初始化占用UI线程
                var t = new Timer { Interval = 500 };
                t.Tick += (_, __) => { t.Stop(); t.Dispose(); Start(); };
                t.Start();
            }
        }

        protected override void ExitThreadCore()
        {
            try { Stop(); } catch { }
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
            base.ExitThreadCore();
        }

        private void ShowSettings()
        {
            using var dlg = new SettingsForm();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (_engine is not null)
                {
                    // 重启以应用新参数
                    Stop();
                    Start();
                }
            }
        }

        private void Start()
        {
            if (_engine != null) return;
            var cfg = AppSettings.Load();

            try
            {
                _engine = new AudioEngine(cfg);
                _engine.Start();
                _startItem.Enabled = false;
                _stopItem.Enabled = true;
                _tray.Text = "MirrorAudio（运行中）";
            }
            catch (Exception ex)
            {
                _engine?.Dispose();
                _engine = null;
                MessageBox.Show("启动失败：\r\n" + ex.Message, "MirrorAudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Stop()
        {
            if (_engine == null) return;
            try { _engine.Dispose(); } finally { _engine = null; }
            _startItem.Enabled = true;
            _stopItem.Enabled = false;
            _tray.Text = "MirrorAudio（已停止）";
        }
    }
}