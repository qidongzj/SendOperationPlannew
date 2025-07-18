using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendOperationPlan
{
    public partial class FormUpdate : Form
    {
        public FormUpdate()
        {
            InitializeComponent();
        }

        private void FormUpdate_Load(object sender, EventArgs e)
        {
           

            string downloadDir = @"\\172.20.1.35\outpwpf";//正定
            //下载的程序放置的子目录名称
            string childDir = "outp";
            //启动exe名称
            string appName = "Synyi.Outp.Wpf.Launcher.exe";

            //创建自动子目录 
            string appChildDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, childDir);
            if (Directory.Exists(appChildDir) == false)
            {
                Directory.CreateDirectory(appChildDir);
            }

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();

            string xcopy = $@"xcopy {downloadDir}\*.* %CD%\{childDir} /s /d /r";
            process.StandardInput.WriteLineAsync(xcopy);
            process.StandardInput.FlushAsync();
            process.StandardInput.WriteLineAsync($@"cd %CD%\{childDir}");
            process.StandardInput.FlushAsync();
            process.StandardInput.WriteLineAsync($@"start {appName}");
            process.StandardInput.FlushAsync();
            process.WaitForExit(2000);
            Application.Exit();
        }
    }
}
