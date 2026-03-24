using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace WordReplaceLauncher
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                
                // 方案二：继续使用 python.exe
                string pythonExe = Path.Combine(baseDir, "runtime", "python.exe");
                string mainScript = Path.Combine(baseDir, "run.py");

                if (!File.Exists(pythonExe))
                {
                    MessageBox.Show("找不到内置环境 (runtime\\python.exe)。", "错误");
                    return;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = pythonExe;
                startInfo.Arguments = "\"" + mainScript + "\"";
                startInfo.WorkingDirectory = baseDir;
                startInfo.UseShellExecute = false;
                
                // 核心：强制不创建窗口，并隐藏
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;

                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败: " + ex.Message, "错误");
            }
        }
    }
}
