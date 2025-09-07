using System;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace MirrorAudio.AppContextApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // 单实例
            using var mtx = new System.Threading.Mutex(true, "Global\\MirrorAudio.AppContext", out bool created);
            if (!created) return;

            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // ApplicationContext 模式
            using var ctx = new TrayAppContext();
            Application.Run(ctx);
        }
    }
}