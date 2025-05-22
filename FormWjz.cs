using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace SendOperationPlan
{
    /// <summary>
    /// 危急值推送窗体
    /// </summary>
    public partial class FormWjz : Form
    {
        private static string corpid = "wx908901486f530786";            // 企业微信后台获取
        private static string corpsecret = "jOpxoJf1Z2HI5IuxCnGURQQTOiwLSjylQJW7YIsBfj4";    // 对应应用的密钥
        private static string tokenUrl = $"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpid}&corpsecret={corpsecret}";
        private static string sendUrl = $"https://qyapi.weixin.qq.com/cgi-bin/message/send";  // 发送消息的URL

        private bool isEnabled = true;

        public FormWjz()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 获取token
        /// </summary>
        /// <returns></returns>
        public string GetAccessToken()
        {
            using (HttpClient client = new HttpClient())
            {
                var response = client.GetAsync(tokenUrl).Result;
                response.EnsureSuccessStatusCode();
                dynamic result = JsonConvert.DeserializeObject(response.Content.ReadAsStringAsync().Result);
                return result.access_token;  // 有效期7200秒，需缓存复用[3](@ref)
            }
        }

        public void SendTextMessage(string apiUrl, string accessToken, string userid, string content1)
        {
            //string apiUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send";
            var message = new
            {
                touser = userid,
                msgtype = "text",
                agentid = "1000024",
                text = new { content = content1 }
            };

            using (HttpClient client = new HttpClient())
            {
                var content2 = new StringContent(JsonConvert.SerializeObject(message), Encoding.UTF8, "application/json");
                var response = client.PostAsync($"{apiUrl}?access_token={accessToken}", content2).Result;
                // Console.WriteLine(response.Content.ReadAsStringAsync().Result);  // 输出接口返回结果[4](@ref)
                richTextBox1.Text = response.Content.ReadAsStringAsync().Result;
            }
        }

        public string SendTextMessageMarkdown(string apiUrl, string accessToken, string userid, string content1)
        {
            //string apiUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send";
            try
            {
                var message = new
                {
                    touser = userid,
                    msgtype = "markdown",
                    agentid = "1000024",
                    markdown = new { content = content1 }
                };

                using (HttpClient client = new HttpClient())
                {
                    var content2 = new StringContent(JsonConvert.SerializeObject(message), Encoding.UTF8, "application/json");
                    var response = client.PostAsync($"{apiUrl}?access_token={accessToken}", content2).Result;
                    // Console.WriteLine(response.Content.ReadAsStringAsync().Result);  // 输出接口返回结果[4](@ref)
                    //richTextBox1.Text = response.Content.ReadAsStringAsync().Result;
                    return response.Content.ReadAsStringAsync().Result;
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog("企业微信接口异常:" + ex.Message);
                return ex.Message;
            }
        }

        private void FormWjz_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) // 仅处理用户点击关闭按钮
            {
                e.Cancel = true;        // 取消关闭操作
                this.WindowState = FormWindowState.Minimized; // 最小化窗口
                this.ShowInTaskbar = false;     // 隐藏任务栏图标
                this.notifyIcon1.Visible = true;      // 显示托盘图标
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.ShowInTaskbar = true;          // 显示任务栏图标
            this.WindowState = FormWindowState.Normal; // 恢复窗口
            this.Activate();                    // 激活窗口到前台
            this.notifyIcon1.Visible = false;         // 隐藏托盘图标
        }

        private void 最大化ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal; // 恢复窗口
            this.Activate();                    // 激活窗口到前台
            notifyIcon1.Visible = false;
        }

        private void FormWjz_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
              "确认退出程序吗？  \r\n 本程序是推送危急值消息推送到医生企业微信的实时通知！ 如您不清楚情况，请不要退出本程序，请点 “否” 按钮 ！",
              "危急值实时推送计划",
              MessageBoxButtons.YesNo,
              MessageBoxIcon.Question
          );
            if (result == DialogResult.Yes)
            {
                this.timer1?.Dispose();
                notifyIcon1.Visible = false; // 隐藏托盘图标
                Application.Exit();         // 彻底退出程序
            }
        }

        private void FormWjz_Load(object sender, EventArgs e)
        {
            isEnabled = true;

            this.checkBox1.Checked = true;
            this.checkBox2.Checked = true;
            this.timer1.Interval = 60 * 1000; // 60秒轮询一次

            this.timer1?.Start();

            button3.Enabled = false;
            button4.Enabled = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 获取路径（推荐AppDomain方式）
            string path = AppDomain.CurrentDomain.BaseDirectory + "log";

            // 验证路径是否存在
            if (Directory.Exists(path))
            {
                // 打开资源管理器
                Process.Start("explorer.exe", path);
                //Console.WriteLine("路径已打开: " + path);
            }
            else
            {
                //Console.WriteLine("路径不存在！");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.timer1?.Stop();
            //this.timer2?.Stop();
            button3.Enabled = true;
            button4.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.timer1?.Start();
            //this.timer2?.Start();
            button3.Enabled = false;
            button4.Enabled = true;
        }


        /// <summary>
        /// 记录日志
        /// </summary>
        /// <param name="log"></param>
        private void WriteErrorLog(string log)
        {
            //if (!LogEnabled) return;

            if (!Directory.Exists(Path.Combine(Application.StartupPath, "log"))) return;
            try
            {
                string logfile = Path.Combine(Application.StartupPath, "log\\ErrorMessage_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                FileStream fs = new FileStream(logfile, FileMode.Append);
                try
                {
                    StreamWriter sw = new StreamWriter(fs);
                    try
                    {
                        sw.Write(string.Format("{0}\r\n{1}\r\n{2}\r\n",
                                                   "--------------------------------------------------",
                                                   DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                   log
                                                   )
                                     );
                    }
                    finally
                    {
                        sw.Close();
                    }
                }
                finally
                {
                    fs.Close();
                }
            }
            catch
            { }
        }

        /// <summary>
        /// 记录日志
        /// </summary>
        /// <param name="log"></param>
        private void WriteLog(string log)
        {
            //if (!LogEnabled) return;

            if (!Directory.Exists(Path.Combine(Application.StartupPath, "log"))) return;
            try
            {
                string logfile = Path.Combine(Application.StartupPath, "log\\pushWjzMessage_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                FileStream fs = new FileStream(logfile, FileMode.Append);
                try
                {
                    StreamWriter sw = new StreamWriter(fs);
                    try
                    {
                        sw.Write(string.Format("{0}\r\n{1}\r\n{2}\r\n",
                                                   "--------------------------------------------------",
                                                   DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                   log
                                                   )
                                     );
                    }
                    finally
                    {
                        sw.Close();
                    }
                }
                finally
                {
                    fs.Close();
                }
            }
            catch
            { }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 推送危急诊消息 
        /// </summary>
        /// <param name="type">1=门诊 ，2住院</param>
        private void SendMessage(int type, bool istest)
        {

        }
    }
}
