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
using DocumentFormat.OpenXml.EMMA;
using Newtonsoft.Json;
using static ClosedXML.Excel.XLPredefinedFormat;
using DateTime = System.DateTime;

namespace SendOperationPlan
{
    public partial class FormHospitalReadmission : Form
    {
        private static string corpid = "wx908901486f530786";            // 企业微信后台获取
        private static string corpsecret = "jOpxoJf1Z2HI5IuxCnGURQQTOiwLSjylQJW7YIsBfj4";    // 对应应用的密钥
        private static string tokenUrl = $"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpid}&corpsecret={corpsecret}";
        private static string sendUrl = $"https://qyapi.weixin.qq.com/cgi-bin/message/send";  // 发送消息的URL
        private System.Timers.Timer _timer2; // 每100秒检查

        private System.Timers.Timer _timer3;//每90秒检查一次

        //七日重复入院使用
        private bool IsEnable3 = false;

        //外国人入院使用
        private bool IsEnable4  = false;

        public FormHospitalReadmission()
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
                string logfile = Path.Combine(Application.StartupPath, "log\\pushMessage_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
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
        /// 
        /// </summary>
        /// <param name="log"></param>
        

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

        private void button3_Click(object sender, EventArgs e)
        {
            _timer2.Interval = 100000; // 每100秒检查
            _timer2?.Start();

            _timer3.Interval=90000; // 每90秒检查一次
            _timer3?.Start();




            this.IsEnable3 = false; // 重置标志位
            this.IsEnable4 = false;

            button3.Enabled = false;
            button4.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            _timer2?.Stop();

            _timer3?.Stop();

            button3.Enabled = true;
            button4.Enabled = false;

            this.IsEnable3 = false; // 重置标志位
            this.IsEnable4 = false;
        }

        private void FormHospitalReadmission_Load(object sender, EventArgs e)
        {
            this._timer2 = new System.Timers.Timer(100 * 1000); // 每100秒检查
            this._timer2?.Start();
            this._timer3 = new System.Timers.Timer(90 * 1000); // 每100秒检查
            this._timer3?.Start();


            _timer2.Elapsed += (s, e1) =>
            {
                if (IsEnable3 || (e1.SignalTime.Hour == 7 && (e1.SignalTime.Minute == 0 || e1.SignalTime.Minute == 1)))
                {


                    if (checkBox2.Checked)
                    {
                        //早上7点推送给管理员 高洁 和我 ， (昨天入院的患者,前面7日内出院的)
                        SendHospitalReadmissionMessage(false);
                    }
                   
                    if (this._timer2.Interval == 100 * 1000)
                    {
                        _timer2.Stop(); // 停止定时器
                        _timer2.Interval = TimeSpan.FromHours(24).TotalMilliseconds; // 重置为24小时
                        _timer2.Start(); // 重新启动定时器
                        IsEnable3 = true;
                    }
                }
            };


      
            _timer3.Elapsed += (s, e1) =>
            {
                if (IsEnable4 ||  (e1.SignalTime.Minute == 0 || e1.SignalTime.Minute == 1))
                {
                    if (checkBox1.Checked)
                    {
                        //每小时推送给管理员 高洁 和我 
                        SendForeginMessage(false);

                    
                    }

                    if (this._timer3.Interval == 90 * 1000)
                    {
                        _timer3.Stop(); // 停止定时器
                        _timer3.Interval = TimeSpan.FromHours(1).TotalMilliseconds; // 重置为1小时
                        _timer3.Start(); // 重新启动定时器
                        IsEnable4 = true;
                    }
                }
            };


            //comboBox1.SelectedIndex = 0;
            button3.Enabled = false;
            button4.Enabled = true;
        }



        // 是否在7天内
        public bool IsPaidWithin7Days(DateTime orderTime, DateTime paymentTime)
        {
            TimeSpan diff = orderTime > paymentTime
                           ? orderTime - paymentTime
                           : paymentTime - orderTime;
            return diff.TotalSeconds <= 604800;
        }


        // bool isValid = IsPaidWithin7Days(order, payment); // true（相差 518,399 秒）

        /// <summary>
        /// 外国人门诊住院消息推送
        /// </summary>
        /// <param name="istest"></param>
        private void SendForeginMessage(bool istest)
        {

            int day = 1;
            if (istest)
            {
                day = 1;
            }
            string dt = DateTime.Now.AddDays(-day).ToString("yyyy-MM-dd").Replace("-","");

            // string sql2  = "select  '门诊' MzOrZy ,FLOOR(DATEDIFF(DAY, birth, GETDATE()) / 365.25) AS age,lrrq as dyrq,sex, patid,blh,hzxm,cardno,sfzh  from THIS4.dbo.SF_BRXXK WHERE zjlx='26' and lrrq>'" + dt + "' \r\n union\r\n select  '住院' MzOrZy,FLOOR(DATEDIFF(DAY, birth, GETDATE()) / 365.25) AS age,djrq as dyrq,sex,  patid,blh,hzxm,cardno,sfzh  from THIS4.dbo.ZY_BRXXK WHERE  zjlx='26' and djrq>'" + dt+"'";
            string sql = string.Empty;

            if (istest)
            {
                sql = "select  '门诊' MzOrZy,FLOOR(DATEDIFF(DAY, birth, GETDATE()) / 365.25) AS age ,case  a.zjlx when 1 then '护照' when 26 then '外国人永久居留身份证' else '其他' end as zjlx,b.ghrq as dyrq,sex, a.patid,a.blh,a.hzxm,a.sfzh,b.ksmc ,c.lbmc as ghlbmc ,d.rqflmc ybmc   from THIS4.dbo.SF_BRXXK a (nolock) join THIS4.dbo.GH_GHZDK b (nolock)\r\non a.patid=b.patid left join THIS4.dbo.GH_GHLBK c on c.lbxh=b.ghlb left join THIS4.dbo.YY_YBFLK d on d.ybdm=b.ybdm \r\nWHERE a.zjlx in ('1','9','21','26') \r\nand len(a.sfzh)!=18 and len(a.sfzh)>0  and b.ghrq>'" + dt + "'\r\n union\r\n select  '住院' MzOrZy,FLOOR(DATEDIFF(DAY, a.birth, GETDATE()) / 365.25) AS age,case  a.zjlx when 1 then '护照' when 26 then '外国人永久居留身份证' else '其他' end as zjlx, b.ryrq as dyrq, a.sex, a.patid,a.blh,a.hzxm,a.sfzh,c.name as ksmc , '' as ghlbmc,d.rqflmc ybmc from THIS4.dbo.ZY_BRXXK a (nolock) join THIS4.dbo.ZY_BRSYK b(nolock) \r\n on a.patid=b.patid \r\n join THIS4.dbo.YY_KSBMK c(nolock) on c.id=b.ksdm left join THIS4.dbo.YY_YBFLK d on d.ybdm=b.ybdm \r\n  WHERE  a.zjlx in ('1','9','21','26') and len(a.sfzh)>0   and len(a.sfzh)!=18 and b.ryrq>'" + dt + "' and a.sfzh not like '未报%' and a.sfzh not like '户口%' and a.sfzh not like '新生%'  and a.sfzh not like '未上户口%'\r\n  and a.sfzh not like '未带%'   and a.sfzh not like '儿童%'   and a.sfzh not like '出生儿未上户口%'  and a.sfzh not like '无户口%' and a.sfzh not like '没有%' and a.sfzh not like '无%'";
            }
            else 
            {
                DateTime now = DateTime.Now;
                DateTime hourZero = now.Date.AddHours(now.Hour); // 精确到小时，分秒归零
                string d1 = hourZero.ToString("yyyyMMddHH:mm:ss");
                string d2 = hourZero.AddHours(-1).ToString("yyyyMMddHH:mm:ss");

                sql = "select  '门诊' MzOrZy,FLOOR(DATEDIFF(DAY, birth, GETDATE()) / 365.25) AS age ,case  a.zjlx when 1 then '护照' when 26 then '外国人永久居留身份证' else '其他' end as zjlx,b.ghrq as dyrq,sex, a.patid,a.blh,a.hzxm,a.sfzh,b.ksmc ,c.lbmc as ghlbmc ,d.rqflmc ybmc   from THIS4.dbo.SF_BRXXK a (nolock) join THIS4.dbo.GH_GHZDK b (nolock)\r\non a.patid=b.patid left join THIS4.dbo.GH_GHLBK c on c.lbxh=b.ghlb left join THIS4.dbo.YY_YBFLK d on d.ybdm=b.ybdm \r\nWHERE a.zjlx in ('1','9','21','26') \r\nand len(a.sfzh)!=18 and len(a.sfzh)>0  and ( b.ghrq<'"+d1+"' and b.ghrq>='" + d2 + "') \r\n union\r\n select  '住院' MzOrZy,FLOOR(DATEDIFF(DAY, a.birth, GETDATE()) / 365.25) AS age,case  a.zjlx when 1 then '护照' when 26 then '外国人永久居留身份证' else '其他' end as zjlx, b.ryrq as dyrq, a.sex, a.patid,a.blh,a.hzxm,a.sfzh,c.name as ksmc , '' as ghlbmc,d.rqflmc ybmc from THIS4.dbo.ZY_BRXXK a (nolock) join THIS4.dbo.ZY_BRSYK b(nolock) \r\n on a.patid=b.patid \r\n join THIS4.dbo.YY_KSBMK c(nolock) on c.id=b.ksdm left join THIS4.dbo.YY_YBFLK d on d.ybdm=b.ybdm \r\n  WHERE  a.zjlx in ('1','9','21','26') and len(a.sfzh)>0   and len(a.sfzh)!=18 and ( b.ryrq<'"+d1+"' and  b.ryrq>='" + d2 + "') and a.sfzh not like '未报%' and a.sfzh not like '户口%' and a.sfzh not like '新生%'  and a.sfzh not like '未上户口%'\r\n  and a.sfzh not like '未带%'   and a.sfzh not like '儿童%'   and a.sfzh not like '出生儿未上户口%'  and a.sfzh not like '无户口%' and a.sfzh not like '没有%' and a.sfzh not like '无%'";

            }

            DataTable dt2 = DbHelper.GetData(sql, CommandType.Text, null);
            if (dt2 == null || dt2.Rows.Count == 0)
            {
                WriteLog("当下无外籍人士门诊住院就诊");
                return;
            }
            else
            {
                List<ForeginInfo> foreginlist = new List<ForeginInfo>();
                foreach (DataRow dr in dt2.Rows) 
                {
                    ForeginInfo foregin = new ForeginInfo();
                    foregin.MzOrZy = dr["MzOrZy"].ToString();
                    foregin.patid = dr["patid"].ToString();
                    foregin.blh = dr["blh"].ToString();
                    foregin.hzxm = dr["hzxm"].ToString();
                    //foregin.cardno = dr["cardno"].ToString();
                    foregin.sfzh = dr["sfzh"].ToString();
                    foregin.age=dr["age"].ToString();
                    foregin.sex = dr["sex"].ToString();
                    foregin.dyrq = dr["dyrq"].ToString(); // 到院日期
                    foregin.ksmc = dr["ksmc"].ToString() ;
                    foregin.zjlx = dr["zjlx"].ToString(); // 就诊类型
                    foregin.ghlbmc = dr["ghlbmc"].ToString(); 
                    foregin.ybmc = dr["ybmc"].ToString(); // 医保名称
                    foreginlist.Add(foregin);
                }

                if (foreginlist.Count > 0)
                {
                    string accessToken = GetAccessToken();
                    if (string.IsNullOrEmpty(accessToken))
                    {
                        WriteErrorLog("获取企业微信AccessToken失败");
                        return;
                    }
                    //string content = "";
                    string userid = "2382|3024"; // 
                    if (istest)
                    {
                        userid = "2382|3024"; // 测试账号
                    }

                    string datetime2 = DateTime.Now.ToString("yyyy年MM月dd日");

                    string sendtext = string.Empty;
                    sendtext = $"`外籍人士入院就诊患者信息` \r\n" +
                               $"**事项详情:**  \r\n" +
                               $"<font color=\"info\"> \r\n";
                    int count = 1;
                    foreach (ForeginInfo info in foreginlist)
                    {
                        sendtext += $"【" + (count++) + "】" + info.MzOrZy + "：" +info.ksmc+","+ info.hzxm + "," + info.sex.Trim() + "," + info.age + "岁\r\n";
                        //sendtext += $"卡号:" + info.cardno+"\r\n";
                        sendtext += $"挂号类别:" + info.ghlbmc + "\r\n";
                        sendtext += $"医保费别:" + info.ybmc + "\r\n";
                        sendtext += $"证件类型:" + info.zjlx + "\r\n";
                        sendtext += $"证件号:" +info.sfzh + "\r\n";
                        if (info.MzOrZy == "住院")
                        {
                            sendtext += $"入院日期:" + info.dyrq + "\r\n";
                        }
                        else 
                        {
                            sendtext += $"就诊日期:" + info.dyrq + "\r\n";
                        }
                    }
                    sendtext += $"</font>  \r\n";

                    sendtext += $"推送日期：<font color=\"warning\">{datetime2}</font>  \n";
                    sendtext += $" \r\n";

                    WriteLog($"推送入参: 工号： {userid}, 入参 {sendtext}");
                    string result = SendTextMessageMarkdown(sendUrl, accessToken, userid, sendtext);
                    WriteLog($"推送结果: {result}");

                    //string sendtext = string.Empty;
                    //sendtext += $"`患者七日内重复入院的列表`\r\n";
                    //sendtext += $"**事项详情:**  \r\n";
                    //sendtext += $"<font color=\"info\"> \r\n";
                    //int count = 1;
                    //foreach (Result info in resultlist)
                    //{

                    //    sendtext += $"【" + (count++) + "】" + info.HZXM + "," + info.SEX.Trim() + "," + info.XSNL + "," + info.YBMC + "  \r\n";
                    //    sendtext += $"上次出院:" + info.SCZYH + "," + info.CYKSMC + "," + info.CYRQ.ToString("D") + " " + info.CYRQ.Hour + "时, " + info.CYZDMC + ", 住院天数:" + info.ZYTS + "  \r\n";
                    //    sendtext += $"本次入院:" + info.BCZYH + "," + info.RYKSMC + "," + info.RYRQ.ToString("D") + " " + info.RYRQ.Hour + "时, " + info.RYZDMC + "  \r\n";
                    //    sendtext += "\r\n";
                    //}
                    //sendtext += $"</font>  \r\n";
                    //sendtext += $"推送日期：<font color=\"warning\">{datetime2}</font>  \n";
                    //sendtext += $" \r\n";







                }

            }
        }

        /// <summary>
        /// (昨天入院的患者,前面7日内出院的)
        /// </summary>
        /// <param name="istest"></param>

        private void SendHospitalReadmissionMessage(bool istest)
        {

            int abc = 0;
            if (istest) 
            {
                abc = comboBox1.SelectedItem == null ? 0 : Convert.ToInt32(comboBox1.SelectedItem);
            }

            //string sql = "select *  from CISDB.dbo.INP_BRSYK a(nolock) where RYRQ>= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-"+(8+abc) +"), 0)  and  RYRQ < DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE() -"+ (0+abc)+"), 0)  and BRZT !=9 and PATID in\r\n(select PATID from CISDB.dbo.INP_BRSYK a(nolock) where  RYRQ>=DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-" + (8 + abc) + "), 0) and  RYRQ < DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-" + (0 + abc) + "), 0)   and BRZT !=9  group by PATID  HAVING COUNT(*) > 1) \r\norder by PATID,a.RYRQ";
            //DataTable dt = DbHelper.GetData(sql, CommandType.Text, null);
            //if (dt == null || dt.Rows.Count == 0)
            //{
            //    WriteLog("没有七日重复办理出入院的患者");
            //    return;
            //}
            //else
            //{
                //List<string> patidlist = new List<string>();
                //foreach (DataRow row in dt.Rows)
                //{
                //    if (!patidlist.Contains(row["PATID"].ToString()))
                //    {
                //        patidlist.Add(row["PATID"].ToString());
                //    }
                //}



                string sql2= "select   *  from CISDB.dbo.INP_BRSYK a(nolock) where RYRQ>= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-" + (1 + abc) + "), 0)  and  RYRQ < DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-" + (0 + abc) + "), 0) and BQDM!='124'  and BRZT !=9  order by PATID,a.RYRQ"; //and PATID in ('" + string.Join("','",patidlist) + "')
                DataTable dt2 = DbHelper.GetData(sql2, CommandType.Text, null);
                if (dt2 == null || dt2.Rows.Count == 0)
                {
                    WriteLog("没有七日重复办理出入院的患者1");
                    return;
                }
                else 
                {

                    //
                    List<string> patidlist2 = new List<string>();
                    foreach (DataRow row in dt2.Rows)
                    {
                        if (!patidlist2.Contains(row["PATID"].ToString()))
                        {
                            patidlist2.Add(row["PATID"].ToString());
                        }
                    }
                    string sql3 = "select *  from CISDB.dbo.INP_BRSYK a(nolock) where CYRQ>= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-" + (8 + abc) + "), 0)  and  CYRQ < DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()-"+ (1 + abc) + "), 0)   and BRZT !=9 and BQDM!='124' and PATID in ('" + string.Join("','", patidlist2) + "') order by PATID asc,a.RYRQ desc";

                    DataTable dt3 = DbHelper.GetData(sql3, CommandType.Text, null);



                    List<INP_BRSYK> daylist  = new List<INP_BRSYK>();
                    foreach (DataRow row in dt2.Rows)
                    {
                        INP_BRSYK brsyk = new INP_BRSYK();
                        brsyk.PATID = row["PATID"].ToString();


                        // Fix for CS8957: Explicitly cast null to DateTime? to ensure compatibility with nullable types.
                        brsyk.RYRQ =  !string.IsNullOrEmpty(row["RYRQ"]?.ToString())  ? Convert.ToDateTime(row["RYRQ"]) : (DateTime?)null;
                        brsyk.CYRQ = !string.IsNullOrEmpty(row["CYRQ"]?.ToString()) ? Convert.ToDateTime(row["CYRQ"]) : (DateTime?)null;
                        brsyk.RQRQ = !string.IsNullOrEmpty(row["RQRQ"]?.ToString()) ? Convert.ToDateTime(row["RQRQ"]) : (DateTime?)null;
                        brsyk.CQRQ = !string.IsNullOrEmpty(row["CQRQ"]?.ToString()) ? Convert.ToDateTime(row["CQRQ"]) : (DateTime?)null;
                        //brsyk.RYRQ = row["CYRQ"] != null ? Convert.ToDateTime(row["RYRQ"]) :null ;
                        //brsyk.CYRQ = Convert.ToDateTime(row["CYRQ"] ?? null);
                        //  brsyk.RQRQ = Convert.ToDateTime(row["RQRQ"] ?? null);
                        // brsyk.CQRQ = Convert.ToDateTime(row["CQRQ"] ?? null);
                        brsyk.KSDM = row["KSDM"].ToString();
                        brsyk.KSMC = row["KSMC"].ToString();
                        brsyk.BQDM = row["BQDM"].ToString();
                        brsyk.BQMC = row["BQMC"].ToString();
                        brsyk.YSDM = row["YSDM"].ToString();
                        brsyk.YSXM = row["YSXM"].ToString();    
                        brsyk.BLH = row["BLH"].ToString();
                        brsyk.HZXM = row["HZXM"].ToString();
                        brsyk.CISXH = decimal.Parse(row["CISXH"].ToString());
                        brsyk.SEX = row["SEX"].ToString();
                        brsyk.XSNL = row["XSNL"].ToString();
                        brsyk.SFZH = row["SFZH"].ToString();
                        brsyk.YBMC = row["YBMC"].ToString();
                        brsyk.RYZDMC = row["RYZDMC"].ToString();
                        brsyk.CYZDMC = row["CYZDMC"].ToString();


                        brsyk.BLH=row["BLH"].ToString();
                        daylist.Add(brsyk);
                    }

                    List<INP_BRSYK> daysevenlist  = new List<INP_BRSYK>();
                    foreach (DataRow row in dt3.Rows)
                    {
                        INP_BRSYK brsyk = new INP_BRSYK();
                        brsyk.PATID = row["PATID"].ToString();
                    brsyk.RYRQ = !string.IsNullOrEmpty(row["RYRQ"]?.ToString()) ? Convert.ToDateTime(row["RYRQ"]) : (DateTime?)null;
                    brsyk.CYRQ = !string.IsNullOrEmpty(row["CYRQ"]?.ToString()) ? Convert.ToDateTime(row["CYRQ"]) : (DateTime?)null;
                    brsyk.RQRQ = !string.IsNullOrEmpty(row["RQRQ"]?.ToString()) ? Convert.ToDateTime(row["RQRQ"]) : (DateTime?)null;
                    brsyk.CQRQ = !string.IsNullOrEmpty(row["CQRQ"]?.ToString()) ? Convert.ToDateTime(row["CQRQ"]) : (DateTime?)null;
                    //brsyk.RYRQ = Convert.ToDateTime(row["RYRQ"] ?? null);
                    //brsyk.CYRQ = Convert.ToDateTime(row["CYRQ"] ?? null);
                    //brsyk.RQRQ = Convert.ToDateTime(row["RQRQ"] ?? null);
                    //brsyk.CQRQ = Convert.ToDateTime(row["CQRQ"] ?? null);
                    brsyk.KSDM = row["KSDM"].ToString();
                        brsyk.KSMC = row["KSMC"].ToString();
                        brsyk.BQDM = row["BQDM"].ToString();
                        brsyk.BQMC = row["BQMC"].ToString();
                        brsyk.YSDM = row["YSDM"].ToString();
                        brsyk.YSXM = row["YSXM"].ToString();
                        brsyk.BLH = row["BLH"].ToString();
                        brsyk.HZXM = row["HZXM"].ToString();
                        brsyk.CISXH = decimal.Parse(row["CISXH"].ToString());
                        brsyk.SEX = row["SEX"].ToString();
                        brsyk.XSNL = row["XSNL"].ToString();
                        brsyk.SFZH = row["SFZH"].ToString();
                        brsyk.YBMC = row["YBMC"].ToString();
                    //brsyk.ZDDM = row["ZDDM"].ToString();
                    //brsyk.ZDMC = row["ZDMC"].ToString();
                        brsyk.RYZDMC = row["RYZDMC"].ToString();
                        brsyk.CYZDMC = row["CYZDMC"].ToString();
                        brsyk.BLH = row["BLH"].ToString();
                    
                        daysevenlist.Add(brsyk);
                    }

                    List<Result> resultlist  = new List<Result>();

                    foreach (INP_BRSYK info in daylist.Where(r=>daysevenlist.Exists(p=>p.PATID== r.PATID)))
                    {
                        INP_BRSYK info2 = daysevenlist.FindAll(x => x.PATID == info.PATID).OrderByDescending(r=>r.CISXH).FirstOrDefault();
                        if (info2 != null)
                        {
                            //如果有昨天入院的患者，且前面7日内出院的
                            if (info.RYRQ > info2.CYRQ   && IsPaidWithin7Days((DateTime)info.RYRQ, (DateTime)info2.CYRQ))
                            {
                                //WriterepeatLog($"重复办理出入院的患者: {info.HZXM}，入院时间: {info.RYRQ}, 出院时间: {info2.CYRQ}");
                                
                                resultlist.Add(new Result
                                {
                                    HZXM = info.HZXM,
                                    XSNL = info.XSNL,
                                    SEX = info.SEX,
                                    CYRQ = (DateTime)info2.CYRQ,
                                    RYRQ = (DateTime)info.RYRQ,
                                    SFZH = info.SFZH,
                                    YBMC = info.YBMC,
                                    CYKSMC=info2.KSMC,
                                    RYKSMC=info.KSMC,
                                    SCZYH=info2.BLH,
                                    CYZDMC = info2.CYZDMC,
                                    BCZYH =info.BLH,
                                    RYZDMC = info.RYZDMC,
                                    ZYTS = Math.Abs((info2.CYRQ.Value.Date - info2.RYRQ.Value.Date).Days),// (info2.CYRQ - info2.RYRQ) + 1, // 住院天数
                                });

                            }
                        }

                    }

                if (resultlist.Count > 0)
                {
                    string accessToken = GetAccessToken();
                    if (string.IsNullOrEmpty(accessToken))
                    {
                        WriteErrorLog("获取企业微信AccessToken失败");
                        return;
                    }
                    //string content = "";
                    string userid = "1774|2382|3024"; // 
                    if (istest)
                    {
                        userid = "3024|2382"; // 测试账号
                    }

                    string datetime2 = DateTime.Now.ToString("yyyy年MM月dd日");

                    //string sendtext = string.Empty;
                    //sendtext = $"`患者七日内重复入院的列表`\r\n" +
                    //           $"**事项详情:**  \r\n" +
                    //           $"<font color=\"info\"> \r\n";
                    //int count = 1;
                    //foreach (Result info in resultlist)
                    //{
                    //    sendtext += $"【" + (count++) + "】" + info.HZXM + "," + info.SEX.Trim() + "," + info.XSNL +","+info.YBMC+ "  \r\n";
                    //    sendtext += $"上次出院:"+info.SCZYH+","+info.CYKSMC+"," + info.CYRQ.ToString("D")+""+ info.CYRQ.Hour+"时, "+ info.CYZDMC+ ", 住院天数:"+info.ZYTS + "  \r\n";
                    //    sendtext += $"本次入院:" + info.BCZYH + "," + info.RYKSMC + "," + info.RYRQ.ToString("D") + "" + info.RYRQ.Hour + "时, " + info.RYZDMC  + "  \r\n";
                    //    sendtext += "\r\n";
                    //}

                    //sendtext += $"</font>";
                    //sendtext += $" \r\n";
                    //sendtext += $"推送日期：<font color=\"warning\">{datetime2}</font>  \n";

                    //sendtext += $" \r\n";



                    string sendtext = string.Empty;
                    sendtext += $"`患者七日内重复入院的列表`\r\n";
                    sendtext += $"**事项详情:**  \r\n";
                    sendtext += $"<font color=\"info\"> \r\n";
                    int count = 1;
                    foreach (Result info in resultlist)
                    {

                        sendtext += $"【" + (count++) + "】" + info.HZXM + "," + info.SEX.Trim() + "," + info.XSNL + "," + info.YBMC + "  \r\n";
                        sendtext += $"上次出院:" + info.SCZYH + "," + info.CYKSMC + "," + info.CYRQ.ToString("D") + " " + info.CYRQ.Hour + "时, " + info.CYZDMC + ", 住院天数:" + info.ZYTS + "  \r\n";
                        sendtext += $"本次入院:" + info.BCZYH + "," + info.RYKSMC + "," + info.RYRQ.ToString("D") + " " + info.RYRQ.Hour + "时, " + info.RYZDMC + "  \r\n";
                        sendtext += "\r\n";
                    }

                    sendtext += $"</font>     \r\n";
                    //sendtext += $" \r\n";
                    sendtext += $"推送日期：<font color=\"warning\">{datetime2}</font>  \n";
                    //sendtext += $"时间：<font color=\"warning\">{infos[0].sstime}</font>  \r\n";
                    sendtext += $" \r\n";
                    //sendtext += $"具体安排情况可能因急诊手术会有调整，请各位医生理解配合!";
                    //sendtext += $" \r\n";





                    WriteLog($"推送入参: 工号： {userid}, 入参 {sendtext}");
                    string result = SendTextMessageMarkdown(sendUrl, accessToken, userid, sendtext);
                    WriteLog($"推送结果: {result}");
                }

                }

            //}

        }

        private void FormHospitalReadmission_FormClosing(object sender, FormClosingEventArgs e)
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

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
      "确认退出程序吗？  \r\n 本程序是推送七日重复入院消息/外籍人士门诊住院就诊消息 推送到医生企业微信的实时通知！ 如您不清楚情况，请不要退出本程序，请点 “否” 按钮 ！",
      "七日重复入院推送计划",
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

        private void FormHospitalReadmission_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendHospitalReadmissionMessage(true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SendForeginMessage(true);
        }

        //private void contextMenuStrip1_Click(object sender, EventArgs e)
        //{
        //    this.ShowInTaskbar = true;
        //    this.WindowState = FormWindowState.Normal; // 恢复窗口
        //    this.Activate();                    // 激活窗口到前台
        //    notifyIcon1.Visible = false;
        //}
    }
}
