using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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

        /// <summary>
        /// 已经发送过的危急值 2天内的
        /// </summary>
        private List<WjzInfo> wjzsendedInfosList   = new List<WjzInfo>();

        /// <summary>
        /// 当前获取可能需要待推送的危急值消息
        /// </summary>
        private List<WjzInfo> wjzBeginSendInfosList  = new List<WjzInfo>();


        /// <summary>
        /// 门诊将要发送的危急值
        /// </summary>
        private List<WjzInfo> wjzInfosmz    = new List<WjzInfo>();

        private List<WjzInfo> wjzInfoszy = new List<WjzInfo>();


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

            //this.checkBox1.Checked = true;
            //this.checkBox2.Checked = true;
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
                string logfile = Path.Combine(Application.StartupPath, "log\\ErrorWjzMessage_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
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


        /// <summary>
        /// 推送危急诊消息 
        /// </summary>
        /// <param name="type">1=门诊 ，2住院,3 全部</param>
        private void SendMessage(bool istest)
        {
            string sql = "SELECT * 　FROM　　SendWeChatCriticalValue(nolock) WHERE SENDTIME >GETDATE()-1 ";
            DataTable dataTable = DbHelper.GetData(sql, CommandType.Text, null);
            wjzsendedInfosList = new List<WjzInfo>();
            wjzBeginSendInfosList = new List<WjzInfo>();
            wjzInfosmz = new List<WjzInfo>();
            wjzInfoszy = new List<WjzInfo>();

           

            //门诊危急值数据
            string sqlmz = "SELECT  * FROM  OUTP_BRWJZXXK(NOLOCK) WHERE  BGSJ >= (CONVERT(varchar, CAST(GETDATE()-1 AS DATE), 112) + '00:00:00') ORDER BY BGSJ ASC,XH ASC";
            DataTable dataTablemz = DbHelper.GetData(sqlmz, CommandType.Text, null);
            if (dataTablemz != null && dataTablemz.Rows.Count>0) 
            {

                foreach (DataRow row in dataTablemz.Rows)
                {
                    WjzInfo wjzInfo = new WjzInfo();
                    wjzInfo.CriticalType = 1; //1门诊 2住院
                    wjzInfo.Hzxm = row["HZXM"].ToString();
                    wjzInfo.Bgdh = row["BGDH"]?.ToString();
                    wjzInfo.Bgsj = row["BGSJ"]?.ToString();
                    wjzInfo.CriticalValue = Convert.ToDecimal(row["XH"]?.ToString().Trim());//危急值主表序号
                    wjzInfo.Wjnr = row["WJNR"]?.ToString();
                    wjzInfo.Patid = Convert.ToDecimal((row["PATID"] ?? "-1").ToString().Trim());
                    wjzInfo.Sjkdysdm = row["SJKDYSDM"]?.ToString();
                    wjzInfo.Sjkdysmc= row["SJKDYSMC"]?.ToString();
                    wjzInfo.Sjksdm = row["SJKSDM"]?.ToString();
                    wjzInfo.Sjksmc = row["SJKDMC"]?.ToString();
                    wjzInfo.Sjbqdm = row["SJBQDM"]?.ToString();
                    wjzInfo.Sjbqmc = row["SJBQMC"]?.ToString();
                    wjzInfo.Sjrq = row["SJRQ"]?.ToString();
                    wjzInfo.Ysdfzt = (row["JLZT"] ?? "-1").ToString().Trim();  //0未答复 1已答复
                    wjzInfo.Ysdfnr = row["YSDFNR"]?.ToString();
                    wjzInfo.Jsysdm = row["JSYSDM"]?.ToString();
                    wjzInfo.Jsysmc = row["JSYSMC"]?.ToString();
                    wjzInfo.Ysjssj = row["JSYSSJ"]?.ToString();
                    wjzInfo.SendTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    
                    //wjzInfo.Zyhm = row["ZYHM"]?.ToString();
                  
                    wjzBeginSendInfosList.Add(wjzInfo);
                }

            }



            //住院危急值数据
            string sqlzy = "SELECT  * FROM  INP_BRWJZXXK(NOLOCK) WHERE  BGSJ >= GETDATE()-1 ORDER BY BGSJ ASC,XH ASC";
            DataTable dataTablezy = DbHelper.GetData(sqlzy, CommandType.Text, null);
            if (dataTablezy != null && dataTablezy.Rows.Count > 0)
            {
                foreach (DataRow row in dataTablezy.Rows)
                {
                    WjzInfo wjzInfo = new WjzInfo();
                    wjzInfo.CriticalType = 2; //1门诊 2住院

                    wjzInfo.Hzxm = row["HZXM"].ToString();
                    wjzInfo.Bgdh = row["BGDH"]?.ToString();
                    wjzInfo.Bgsj = row["BGSJ"]?.ToString();
                    wjzInfo.CriticalValue = Convert.ToDecimal(row["XH"]?.ToString().Trim());//危急值主表序号
                    wjzInfo.Wjnr = row["WJNR"]?.ToString();

                    //wjzInfo.Patid = Convert.ToDecimal((row["PATID"] ?? "-1").ToString().Trim());
                    wjzInfo.Zyhm = row["ZYHM"]?.ToString();
                    wjzInfo.Sjkdysdm = row["SJKDYSDM"]?.ToString();
                    wjzInfo.Sjkdysmc = row["SJKDYSMC"]?.ToString();
                    wjzInfo.Sjksdm = row["SJKSDM"]?.ToString();
                    wjzInfo.Sjksmc = row["SJKSMC"]?.ToString();
                    wjzInfo.Sjbqdm = row["SJBQDM"]?.ToString();
                    wjzInfo.Sjbqmc = row["SJBQMC"]?.ToString();
                    wjzInfo.Sjrq = row["SJRQ"]?.ToString();
                    wjzInfo.Ysdfzt = (row["STATUS"] ?? "-1").ToString().Trim();  //0未答复 1已答复
                    wjzInfo.Ysdfnr = row["YSDFNR"]?.ToString();
                    wjzInfo.Jsysdm = row["JSYSDM"]?.ToString();
                    wjzInfo.Jsysmc = row["JSYSMC"]?.ToString();
                    wjzInfo.Ysjssj = row["JSYSSJ"]?.ToString();
                    wjzInfo.Czsj=row["CZSJ"]?.ToString();//处置时间
                    wjzInfo.Dfysmc = row["DFYSMC"]?.ToString();//应该答复医生姓名
                    wjzInfo.Dfysdm = row["DFYSDM"]?.ToString();//应该答复医生代码
                    wjzInfo.SendTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    wjzBeginSendInfosList.Add(wjzInfo);

                }
            }


            if (wjzBeginSendInfosList.Count == 0)
            {
                return;
            }

            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                //WriteLog("没有需要推送的危急值消息！");
                //return;
                //wjzsendedInfosList= new List<WjzInfo>();
            }
            else
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    WjzInfo wjzInfo = new WjzInfo();
                    wjzInfo.CriticalType = Convert.ToInt32(row["CriticalType"].ToString().Trim()); //1门诊 2住院
                    wjzInfo.Hzxm = row["Hzxm"].ToString();
                    wjzInfo.Bgdh = row["BGDH"]?.ToString();
                    wjzInfo.Bgsj = row["BGSJ"]?.ToString();
                    wjzInfo.CriticalValue = Convert.ToDecimal(row["CriticalValue"]?.ToString().Trim());//危急值主表序号
                    wjzInfo.SendTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    //wjzInfo.Wjnr = row["Wjnr"].ToString();
                    if (wjzInfo.CriticalType == 2)
                    {
                        wjzInfo.Zyhm = row["ZYHM"]?.ToString();
                    }
                    else
                    {
                        wjzInfo.Patid = Convert.ToDecimal((row["PATID"] ?? "-1").ToString().Trim());
                    }
                    wjzInfo.Ysdfzt = (row["YSDFZT"] ?? "-1").ToString().Trim();  //0未答复 1已答复
                    wjzsendedInfosList.Add(wjzInfo);
                }
            }






            //门诊
            if (wjzBeginSendInfosList.Count != 0  && wjzBeginSendInfosList.Exists(r => r.CriticalType == 1) )//门诊CriticalType是1
           {
                foreach (WjzInfo wjzInfo in wjzBeginSendInfosList.Where(r=>r.CriticalType==1))
                {
                    if (!wjzsendedInfosList.FindAll(r=>r.CriticalType== wjzInfo.CriticalType).Exists(r =>  r.CriticalValue==wjzInfo.CriticalValue ))
                    {
                        wjzInfo.Gxlx = 1; //新增
                        wjzInfosmz.Add(wjzInfo);
                    }
                    else 
                    {
                        if (wjzsendedInfosList.FindAll(r => r.CriticalType == wjzInfo.CriticalType).Exists(r =>   r.CriticalValue == wjzInfo.CriticalValue && r.Ysdfzt != wjzInfo.Ysdfzt))
                        {
                            wjzInfo.Gxlx = 2; //更新
                            wjzInfosmz.Add(wjzInfo);
                        }
                    }
                }
           }

           //住院
            if (wjzBeginSendInfosList.Count != 0 && wjzBeginSendInfosList.Exists(r => r.CriticalType == 2) )//住院CriticalType是2
            {
                foreach (WjzInfo wjzInfo in wjzBeginSendInfosList.Where(r => r.CriticalType == 2))
                {
                    if (!wjzsendedInfosList.FindAll(r => r.CriticalType == wjzInfo.CriticalType).Exists(r => r.CriticalValue == wjzInfo.CriticalValue))
                    {
                        wjzInfo.Gxlx = 1; //新增
                        wjzInfoszy.Add(wjzInfo);
                    }
                    else
                    {
                        if (wjzsendedInfosList.FindAll(r => r.CriticalType == wjzInfo.CriticalType).Exists(r => r.CriticalValue == wjzInfo.CriticalValue && r.Ysdfzt != wjzInfo.Ysdfzt))
                        {
                            wjzInfo.Gxlx = 2; //更新
                            wjzInfoszy.Add(wjzInfo);
                        }
                    }
                }
            }

            string datetime2 = DateTime.Now.ToString("yyyy年MM月dd日");
            //string datetime3  = DateTime.Now.ToString("yyyy年MM月dd日");
            //门诊
            if (wjzInfosmz.Count != 0)
            {
                // 第一步推送掉数据  第二天R数据 完全没有的先插入,已有的状态更新掉

                foreach (WjzInfo info in wjzInfosmz)
                {
                    string accessToken = GetAccessToken();
                    string content = "";
                    //string content1 = "";
                    //string content2 = "";
                    //string content3 = "";
                    //string content4 = "";
                    if (info.Ysdfzt == "0")
                    {
                        content = $"`门急诊患者危急值通知(未处理)`\r\n" +
                                    $"**事项详情:**  \r\n" +
                                    $"<font color=\"blue\"> \r\n" +

                                   $" 患者姓名：{info.Hzxm} " +
                                    "\r\n" +
                                   //$" 患者PATID：{info.Patid} " +
                                   //"\r\n" +
                                   //$" 报告单号：{info.Bgdh} " +
                                   //"\r\n" +
                                   $" 送检时间：{info.Sjrq} " +
                                   "\r\n" +
                                   $" 报告时间：{info.Bgsj} " +
                                   "\r\n" +
                                   $" 危急值内容：{info.Wjnr.Replace("\n", "").Replace("\r", "")} " +
                                   "\r\n" +
                                   $" 送检科室：{info.Sjksmc} " +
                                   "\r\n" +
                                   $" 开单医生：{info.Sjkdysmc} " +
                                    $"</font>     \r\n" +
                                   $"消息日期：<font color=\"warning\">{datetime2}</font>  \n";
                        ;

                    }
                    if (info.Ysdfzt == "1")
                    {
                        content = $"`门急诊患者危急值通知(已处理)`\r\n" +
                                    $"**事项详情:**  \r\n" +
                                    $"<font color=\"info\"> \r\n" +
                                   $" 患者姓名：{info.Hzxm} " +
                                    "\r\n" +
                                   //$" 患者PATID：{info.Patid} " +
                                   // "\r\n" +
                                   //$" 报告单号：{info.Bgdh} " +
                                   // "\r\n" +
                                   $" 送检时间：{info.Sjrq} " +
                                    "\r\n" +
                                    $" 报告时间：{info.Bgsj} " +
                                     "\r\n" +
                                    $" 危急值内容：{info.Wjnr.Replace("\n", "").Replace("\r", "")} " +
                                    "\r\n" +
                                    $" 送检科室：{info.Sjksmc} " +
                                     "\r\n" +
                                    $" 开单医生：{info.Sjkdysmc} " +
                                     "\r\n" +
                                    $" 处理医生：{info.Jsysmc} " +
                                     "\r\n" +
                                    $" 处理时间：{info.Ysjssj} " +
                                     "\r\n" +
                                    $" 医生答复内容：{info.Ysdfnr} " +
                                     "\r\n" +
                                    $"</font>     \r\n" +
                                    $"消息日期：<font color=\"warning\">{datetime2}</font>  \n";
                    }


                    string token = GetAccessToken();
                    string msg = string.Empty;
                    if (istest)
                    {
                        //测试
                        msg = SendTextMessageMarkdown(sendUrl, token, textBox1.Text.ToString().Trim(), content);
                    }
                    else
                    {
                        //正式
                        msg = SendTextMessageMarkdown(sendUrl, token, info.Sjkdysdm, content);
                    }

                    dynamic data = JsonConvert.DeserializeObject<dynamic>(msg);
                    if (data.errmsg == "ok")
                    {
                        foreach (WjzInfo info2 in wjzInfosmz)
                        {
                            info2.是否推送成功 = "成功";
                            info2.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    {
                        foreach (WjzInfo info2 in wjzInfosmz)
                        {
                            info2.是否推送成功 = "失败";
                            info2.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    string xxx = string.Empty;
                    if (istest)
                    {
                        xxx = textBox1.Text.ToString().Trim() + " (测试数据)";
                    }
                    else
                    {
                        xxx = info.Sjkdysdm;
                    }

                    WriteLog("url:" + sendUrl + "    \r\n  token:" + token + "  \r\n userid:" + xxx + " \r\n  sendmessage:" + content + " \r\n  回参:" + msg);



                }

                try
                {
                    if (wjzInfosmz.Exists(r => r.Gxlx == 1 && r.是否推送成功 == "成功"))//新增
                    {
                        StringBuilder sql33 = new StringBuilder("INSERT INTO SendWeChatCriticalValue (CriticalType, HZXM,BGDH,BGSJ,CriticalValue,SENDTIME,PATID,YSDFZT) VALUES ");
                        foreach (WjzInfo info in wjzInfosmz.Where(r => r.Gxlx == 1 && r.是否推送成功 == "成功"))
                        {
                            sql33.AppendFormat("(1, '{0}','{1}','{2}',{3},'{4}',{5},{6}),", info.Hzxm, info.Bgdh, info.Bgsj, info.CriticalValue, info.SendTime, info.Patid, info.Ysdfzt);
                        }
                        sql33.Length--; // 移除末尾逗号
                        DbHelper.ExecuteNonQuery(sql33.ToString(), null);
                    }

                    if (wjzInfosmz.Exists(r => r.Gxlx == 2 && r.是否推送成功 == "成功"))//更新
                    {
                        string CriticalValuestr = string.Join(",", wjzInfosmz.Where(r => r.Gxlx == 2 && r.是否推送成功 == "成功").Select(x => x.CriticalValue));

                        string sql45 = "update SendWeChatCriticalValue set YSDFZT=1 where YSDFZT=0  and CriticalType=1 and CriticalValue in (" + CriticalValuestr + ")";

                        WriteLog("更新sql语句:" + sql45);

                        DbHelper.ExecuteNonQuery(sql45, null);

                    }
                }
                catch (Exception ex)
                {
                    WriteErrorLog("更新sql语句异常:" + ex.Message);

                }
                //




                //

            }





            //住院
            if (wjzInfoszy.Count != 0)
            {
                // 第一步推送掉数据  第二天R数据 完全没有的先插入,已有的状态更新掉

                foreach (WjzInfo info in wjzInfoszy)
                {
                    string accessToken = GetAccessToken();
                    string content = "";
                    //string content1 = "";
                    //string content2 = "";
                    //string content3 = "";
                    //string content4 = "";
                    if (info.Ysdfzt == "0")
                    {
                        content = $"`住院患者危急值通知(未处理)`\r\n" +
                                    $"**事项详情:**  \r\n" +
                                    $"<font color=\"blue\"> \r\n" +

                                   $" 患者姓名：{info.Hzxm} " +
                                    "\r\n" +
                                   //$" 患者PATID：{info.Patid} " +
                                   //"\r\n" +
                                   //$" 报告单号：{info.Bgdh} " +
                                   //"\r\n" +
                                   //$" 送检时间：{info.Sjrq} " +
                                   //"\r\n" +
                                   $" 报告时间：{info.Bgsj.Replace("-", "")} " +
                                   "\r\n" +
                                   $" 危急值内容：{info.Wjnr.Replace("\n", "").Replace("\r", "")} " +
                                   "\r\n" +
                                   $" 送检科室：{info.Sjksmc} " +
                                   "\r\n" +
                                   $" 接收医生：{info.Dfysmc} " +
                                    $"</font>     \r\n" +
                                   $"消息日期：<font color=\"warning\">{datetime2}</font>  \n";
                        ;

                    }
                    if (info.Ysdfzt == "1")
                    {
                        content = $"`住院患者危急值通知(已处理)`\r\n" +
                                    $"**事项详情:**  \r\n" +
                                    $"<font color=\"info\"> \r\n" +
                                   $" 患者姓名：{info.Hzxm} " +
                                    "\r\n" +
                                   //$" 患者PATID：{info.Patid} " +
                                   // "\r\n" +
                                   //$" 报告单号：{info.Bgdh} " +
                                   // "\r\n" +
                                   //$" 送检时间：{info.Sjrq} " +
                                   // "\r\n" +
                                    $" 报告时间：{info.Bgsj.Replace("-", "")} " +
                                     "\r\n" +
                                    $" 危急值内容：{info.Wjnr.Replace("\n", "").Replace("\r", "")} " +
                                    "\r\n" +
                                    $" 送检科室：{info.Sjksmc} " +
                                     "\r\n" +
                                    $" 接收医生：{info.Dfysmc} " +
                                     "\r\n" +
                                    $" 处理医生：{info.Jsysmc} " +
                                     "\r\n" +
                                    $" 处理时间：{info.Czsj.Replace("-", "")} " +
                                     "\r\n" +
                                    $" 医生答复内容：{info.Ysdfnr} " +
                                     "\r\n" +
                                    $"</font>     \r\n" +
                                    $"消息日期：<font color=\"warning\">{datetime2}</font>  \n";
                    }


                    string token = GetAccessToken();
                    string msg = string.Empty;
                    if (istest)
                    {
                        //测试
                        msg = SendTextMessageMarkdown(sendUrl, token, textBox1.Text.ToString().Trim(), content);
                    }
                    else
                    {
                        //正式
                        msg = SendTextMessageMarkdown(sendUrl, token, info.Sjkdysdm, content);
                    }

                    dynamic data = JsonConvert.DeserializeObject<dynamic>(msg);
                    if (data.errmsg == "ok")
                    {
                        foreach (WjzInfo info2 in wjzInfoszy)
                        {
                            info2.是否推送成功 = "成功";
                            info2.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    {
                        foreach (WjzInfo info2 in wjzInfoszy)
                        {
                            info2.是否推送成功 = "失败";
                            info2.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    string xxx = string.Empty;
                    if (istest)
                    {
                        xxx = textBox1.Text.ToString().Trim() + " (测试数据)";
                    }
                    else
                    {
                        xxx = info.Dfysdm;
                    }

                    WriteLog("url:" + sendUrl + "    \r\n  token:" + token + "  \r\n userid:" + xxx + " \r\n  sendmessage:" + content + " \r\n  回参:" + msg);



                }

                try
                {
                    if (wjzInfoszy.Exists(r => r.Gxlx == 1 && r.是否推送成功 == "成功"))//新增
                    {
                        StringBuilder sql331 = new StringBuilder("INSERT INTO SendWeChatCriticalValue (CriticalType, HZXM,BGDH,BGSJ,CriticalValue,SENDTIME,ZYHM,YSDFZT) VALUES ");
                        foreach (WjzInfo info in wjzInfoszy.Where(r => r.Gxlx == 1 && r.是否推送成功 == "成功"))
                        {
                            sql331.AppendFormat("(2, '{0}','{1}','{2}',{3},'{4}',{5},{6}),", info.Hzxm, info.Bgdh, info.Bgsj, info.CriticalValue, info.SendTime, info.Zyhm, info.Ysdfzt);
                        }
                        sql331.Length--; // 移除末尾逗号
                        DbHelper.ExecuteNonQuery(sql331.ToString(), null);
                    }

                    if (wjzInfoszy.Exists(r => r.Gxlx == 2 && r.是否推送成功 == "成功"))//更新
                    {
                        string CriticalValuestr = string.Join(",", wjzInfoszy.Where(r => r.Gxlx == 2 && r.是否推送成功 == "成功").Select(x => x.CriticalValue));

                        string sql451 = "update SendWeChatCriticalValue set YSDFZT=1 where YSDFZT=0  and CriticalType=2 and CriticalValue in (" + CriticalValuestr + ")";

                        WriteLog("更新sql语句:" + sql451);

                        DbHelper.ExecuteNonQuery(sql451, null);

                    }
                }
                catch (Exception ex)
                {
                    WriteErrorLog("更新sql语句异常:" + ex.Message);

                }


                dataGridView1.DataSource=wjzInfosmz.Union(wjzInfoszy).ToList();

            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
           
            SendMessage( true);
            

            this.button1.Enabled = true;
           
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
            SendMessage(true);
            this.button1.Enabled = true;
        }

    }
}
