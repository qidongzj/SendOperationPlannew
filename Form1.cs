using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// 必须引用命名空间
using System.Net.Http;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Drawing.Text;
using System.IO;
using System.Diagnostics;  // 推荐使用Newtonsoft.Json处理JSON数据[3,5](@ref)

namespace SendOperationPlan
{
    public partial class Form1 : Form
    {

        private static string corpid = "wx908901486f530786";            // 企业微信后台获取
        private static string corpsecret = "jOpxoJf1Z2HI5IuxCnGURQQTOiwLSjylQJW7YIsBfj4";    // 对应应用的密钥
        private static string tokenUrl  = $"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpid}&corpsecret={corpsecret}";
        private static string sendUrl = $"https://qyapi.weixin.qq.com/cgi-bin/message/send";  // 发送消息的URL

        public Form1()
        {
            InitializeComponent();
        }

        private bool isEnabled = true; 

        public List<OperatInfo> operatInfosall  = new List<OperatInfo>(); // 用于存储手术信息的列表


        /// <summary>
        /// 获取token
        /// </summary>
        /// <returns></returns>
        public  string GetAccessToken()
        {
            using (HttpClient client = new HttpClient())
            {
                var response = client.GetAsync(tokenUrl).Result;
                response.EnsureSuccessStatusCode();
                dynamic result = JsonConvert.DeserializeObject(response.Content.ReadAsStringAsync().Result);
                return result.access_token;  // 有效期7200秒，需缓存复用[3](@ref)
            }
        }

        public  void SendTextMessage(string apiUrl,string accessToken, string userid, string content1)
        {
            //string apiUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send";
            var message = new
            {
                touser = userid,
                msgtype = "text",
                agentid = "1000024",  
                text = new { content = content1  }
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
            catch(Exception ex)
            {
                
                return ex.Message;
            }
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
        private void WriterepeatLog(string log)
        {
            //if (!LogEnabled) return;

            if (!Directory.Exists(Path.Combine(Application.StartupPath, "log"))) return;
            try
            {
                string logfile = Path.Combine(Application.StartupPath, "log\\pushrepeatMessage_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
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

        private void button1_Click(object sender, EventArgs e)
        {
            //纯文本
            //SendTextMessage(sendUrl, GetAccessToken(), "3024", "你好啊");

            //富文本
            //SendTextMessageMarkdown(sendUrl, GetAccessToken(), "3024", "`您的手术排班通知`\r\n **事项详情:**  < font color =\"info\"> \r\n (1) 7号楼-2号 /第1台/张三，女，44岁/左股骨骨折固定术/麻醉方式/ 张医生  \r\n (2) 7号楼-2号/第2台/李四，女，87岁/左股骨骨折固定术/麻醉方式/张医生 \r\n (3) 1号楼-3号/第4台/王五，女，88岁/左股骨骨折固定术/麻醉方式/高医生 </font>     \r\n \r\n 日　期：<font color=\"warning\">2025年5月16日</font>  \n 时　间：<font color=\"warning\">上午9:00-11:00</font>  \r\n \r\n 请合理安排时间准时参加手术!");

            this.button1.Enabled = false;
            this.button2.Enabled = false;


            //当天
            SendMessage(false,true);

            this.button1.Enabled = true;
            this.button2.Enabled = true;

        }


        //false 当天 ；true 第二天 
        private void SendMessage(bool T,bool istest)
        {
            
            //考虑日期发生时间 是前一天下午 还是第二天早上

            DateTime dateTime = DateTime.Now;
            if (T)
            {
                dateTime = dateTime.AddDays(1);
            }
            

            //下午三点推送第二天的手术排班
            //上午7点推送当天的手术排班

            string datetime = dateTime.ToString("yyyy-MM-dd");
            string datetime2 = dateTime.ToString("yyyy年MM月dd日");

            /*
            string sql = "SELECT A.*, B.NAME AS PAT_NAME, YEAR(getdate()) - YEAR(B.DATE_OF_BIRTH) AS PAT_AGE, B.INP_NO, B.SEX,ISNULL(rtrim(L.USER_NAME), " +
                "A.FIRST_ASSISTANT) AS FIRST_ASSISTANT_NAME, ISNULL(rtrim(M.USER_NAME), A.SECOND_ASSISTANT) AS SECOND_ASSISTANT_NAME, ISNULL(rtrim(N.USER_NAME), A.THIRD_ASSISTANT) AS THIRD_ASSISTANT_NAME," +
                "ISNULL(D.USER_NAME, A.SURGEON) AS SURGEON_NAME,  E.USER_NAME AS ANESTHESIA_DOCTOR_NAME, F.USER_NAME AS ANESTHESIA_ASSISTANT_NAME,  O.USER_NAME AS ANESTHESIA_ASSISTANT_NAME_TWO," +
                " G.USER_NAME AS FIRST_OPERATION_NURSE_NAME, H.USER_NAME AS SECOND_OPERATION_NURSE_NAME, I.USER_NAME AS FIRST_SUPPLY_NURSE_NAME, J.USER_NAME AS SECOND_SUPPLY_NURSE_NAME,  " +
                " K.USER_NAME AS THIRD_SUPPLY_NURSE_NAME, Q.DEPT_NAME AS WARD_CODE_NAME, case when a.visit_id = 0 then '门诊' else '住院' end as SCHEDULE_TYPE ,   " +
                "    CASE WHEN ISNULL(RTRIM(OPERATING_ROOM_NO), '-1') = '-1' THEN  T.DEPT_NAME  ELSE  T.DEPT_NAME + '(' + CONVERT(VARCHAR, A.SEQUENCE) + ')' END AS OPERATING_DEPT_NAME,  CASE WHEN A.EMERGENCY_INDICATOR = 0 AND" +
                "  A.SCHEDULED_DATE_TIME IS NOT NULL AND              CONVERT(VARCHAR(10), P.ADMISSION_DATE_TIME, 121) =  CONVERT(VARCHAR(10), A.SCHEDULED_DATE_TIME, 121) THEN  1  ELSE   0  END AS DAYTIME_OPERATION,SS.*,MM.BED_LABEL FROM MED_OPERATION_SCHEDULE A" +
                "  LEFT OUTER JOIN MED_PAT_MASTER_INDEX B    ON A.PATIENT_ID = B.PATIENT_ID  LEFT OUTER JOIN MED_HIS_USERS D   ON A.SURGEON = D.USER_ID  LEFT OUTER JOIN MED_HIS_USERS E    ON A.ANESTHESIA_DOCTOR = E.USER_ID  LEFT OUTER JOIN" +
                " MED_HIS_USERS F    ON A.ANESTHESIA_ASSISTANT = F.USER_ID  LEFT OUTER JOIN MED_HIS_USERS G    ON A.FIRST_OPERATION_NURSE = G.USER_ID  LEFT OUTER JOIN" +
                " MED_HIS_USERS H    ON A.SECOND_OPERATION_NURSE = H.USER_ID  LEFT OUTER JOIN MED_HIS_USERS I   ON A.FIRST_SUPPLY_NURSE = I.USER_ID  LEFT OUTER JOIN MED_HIS_USERS J  ON A.SECOND_SUPPLY_NURSE = J.USER_ID  LEFT " +
                "OUTER JOIN MED_HIS_USERS K    ON A.THIRD_SUPPLY_NURSE = K.USER_ID  LEFT OUTER JOIN MED_HIS_USERS L    ON A.FIRST_ASSISTANT = L.USER_ID  LEFT OUTER JOIN MED_HIS_USERS M    ON A.SECOND_ASSISTANT = M.USER_ID" +
                "  LEFT OUTER JOIN MED_HIS_USERS N    ON A.THIRD_ASSISTANT = N.USER_ID  LEFT OUTER JOIN MED_DEPT_DICT T    ON A.OPERATING_DEPT = T.DEPT_CODE  LEFT OUTER JOIN MED_HIS_USERS O    ON A.SECOND_ANESTHESIA_ASSISTANT = O.USER_ID" +
                "  LEFT OUTER JOIN MED_PATS_IN_HOSPITAL P    ON A.PATIENT_ID = P.PATIENT_ID   AND A.VISIT_ID = P.VISIT_ID  LEFT OUTER JOIN MED_DEPT_DICT Q    ON P.WARD_CODE = Q.DEPT_CODE  LEFT OUTER JOIN MED_SCHEDULED_OPERATION_NAME SS ON A.PATIENT_ID =SS.PATIENT_ID " +
                " LEFT OUTER JOIN MED_SCHEDULED_OPERATION_NAME TT on SS.PATIENT_ID =TT.PATIENT_ID and SS.RESERVED1< TT.RESERVED1  left outer join MED_OPERATING_ROOM MM on  CONVERT(NVARCHAR(16), MM.BED_ID)=isnull(A.OPERATING_ROOM_NO,'-80')  WHERE TT.PATIENT_ID is NULL and  (CONVERT(VARCHAR, A.SCHEDULED_DATE_TIME, 23) =  @SheduleDateTime )" +
                "   AND A.OPERATING_ROOM = @OperatingRoom   AND NOT A.OPERATING_ROOM_NO IS NULL   AND A.STATE IN (0, 1, 2) ORDER BY ISNULL((SELECT BED_ID   FROM MED_OPERATING_ROOM   WHERE A.OPERATING_ROOM_NO = ROOM_NO),   0),   A.SEQUENCE";
            */

            string sql = "SELECT A.*, B.NAME AS PAT_NAME, YEAR(getdate()) - YEAR(B.DATE_OF_BIRTH) AS PAT_AGE, B.INP_NO, B.SEX,ISNULL(rtrim(L.USER_NAME), " +
                "A.FIRST_ASSISTANT) AS FIRST_ASSISTANT_NAME, ISNULL(rtrim(M.USER_NAME), A.SECOND_ASSISTANT) AS SECOND_ASSISTANT_NAME, ISNULL(rtrim(N.USER_NAME), A.THIRD_ASSISTANT) AS THIRD_ASSISTANT_NAME," +
                "ISNULL(D.USER_NAME, A.SURGEON) AS SURGEON_NAME,  E.USER_NAME AS ANESTHESIA_DOCTOR_NAME, F.USER_NAME AS ANESTHESIA_ASSISTANT_NAME,  O.USER_NAME AS ANESTHESIA_ASSISTANT_NAME_TWO," +
                " G.USER_NAME AS FIRST_OPERATION_NURSE_NAME, H.USER_NAME AS SECOND_OPERATION_NURSE_NAME, I.USER_NAME AS FIRST_SUPPLY_NURSE_NAME, J.USER_NAME AS SECOND_SUPPLY_NURSE_NAME,  " +
                " K.USER_NAME AS THIRD_SUPPLY_NURSE_NAME, Q.DEPT_NAME AS WARD_CODE_NAME, case when A.VISIT_ID = 0 then '门诊' else '住院' end as SCHEDULE_TYPE ,   " +
                "    CASE WHEN ISNULL(RTRIM(OPERATING_ROOM_NO), '-1') = '-1' THEN  T.DEPT_NAME  ELSE  T.DEPT_NAME + '(' + CONVERT(VARCHAR, A.SEQUENCE) + ')' END AS OPERATING_DEPT_NAME,  CASE WHEN A.EMERGENCY_INDICATOR = 0 AND" +
                "  A.SCHEDULED_DATE_TIME IS NOT NULL AND              CONVERT(VARCHAR(10), P.ADMISSION_DATE_TIME, 121) =  CONVERT(VARCHAR(10), A.SCHEDULED_DATE_TIME, 121) THEN  1  ELSE   0  END AS DAYTIME_OPERATION,SS.*,MM.BED_LABEL FROM [155.155.100.107].docare.dbo.MED_OPERATION_SCHEDULE A" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_PAT_MASTER_INDEX B    ON A.PATIENT_ID = B.PATIENT_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS D   ON A.SURGEON = D.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS E    ON A.ANESTHESIA_DOCTOR = E.USER_ID  LEFT OUTER JOIN" +
                " [155.155.100.107].docare.dbo.MED_HIS_USERS F    ON A.ANESTHESIA_ASSISTANT = F.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS G    ON A.FIRST_OPERATION_NURSE = G.USER_ID  LEFT OUTER JOIN" +
                " [155.155.100.107].docare.dbo.MED_HIS_USERS H    ON A.SECOND_OPERATION_NURSE = H.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS I   ON A.FIRST_SUPPLY_NURSE = I.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS J  ON A.SECOND_SUPPLY_NURSE = J.USER_ID  LEFT " +
                "OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS K    ON A.THIRD_SUPPLY_NURSE = K.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS L    ON A.FIRST_ASSISTANT = L.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS M    ON A.SECOND_ASSISTANT = M.USER_ID" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS N    ON A.THIRD_ASSISTANT = N.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_DEPT_DICT T    ON A.OPERATING_DEPT = T.DEPT_CODE  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS O    ON A.SECOND_ANESTHESIA_ASSISTANT = O.USER_ID" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_PATS_IN_HOSPITAL P    ON A.PATIENT_ID = P.PATIENT_ID   AND A.VISIT_ID = P.VISIT_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_DEPT_DICT Q    ON P.WARD_CODE = Q.DEPT_CODE  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_SCHEDULED_OPERATION_NAME SS ON A.PATIENT_ID =SS.PATIENT_ID " +
                " LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_SCHEDULED_OPERATION_NAME TT on SS.PATIENT_ID =TT.PATIENT_ID and SS.RESERVED1< TT.RESERVED1  left outer join [155.155.100.107].docare.dbo.MED_OPERATING_ROOM MM on  CONVERT(NVARCHAR(16), MM.BED_ID)=isnull(A.OPERATING_ROOM_NO,'-80')  WHERE TT.PATIENT_ID is NULL and  (CONVERT(VARCHAR, A.SCHEDULED_DATE_TIME, 23) =  @SheduleDateTime )" +
                "   AND A.OPERATING_ROOM = @OperatingRoom   AND NOT A.OPERATING_ROOM_NO IS NULL   AND A.STATE IN (0,1,2,3) ORDER BY ISNULL((SELECT BED_ID   FROM [155.155.100.107].docare.dbo.MED_OPERATING_ROOM   WHERE A.OPERATING_ROOM_NO = ROOM_NO),   0),   A.SEQUENCE";




            //住院手术室
            SqlParameter[] parameters = {
            new SqlParameter("@OperatingRoom", "1032"),
            new SqlParameter("@SheduleDateTime", datetime)
            };

            //日间手术室
            SqlParameter[] parameters2 = {
            new SqlParameter("@OperatingRoom", "2020"),
            new SqlParameter("@SheduleDateTime", datetime)
            };
            DataTable dt = DbHelper.GetData(sql, CommandType.Text, parameters);

            DataTable dt2 = DbHelper.GetData(sql, CommandType.Text, parameters2);

            dt.Merge(dt2);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show(datetime2+"没有手术排班记录!");
                return;
            }

            List<OperatInfo> operatInfos = new List<OperatInfo>();
            foreach (DataRow dr in dt.Rows)
            {
                DateTime dtt1 = DateTime.Parse(dr["SCHEDULED_DATE_TIME"].ToString());
                string resulttime = dtt1.ToString("HH:mm");

                OperatInfo info = new OperatInfo()
                {
                    operatingRoomNo = dr["OPERATING_ROOM_NO"].ToString(),
                    operatingRoomNoName = dr["BED_LABEL"].ToString(),
                    sequence = dr["SEQUENCE"].ToString(),

                    sstime = resulttime,
                    wardCodeName = dr["WARD_CODE_NAME"].ToString(),
                    operatdeptName = dr["OPERATING_DEPT_NAME"].ToString(),
                    inpNo = dr["INP_NO"].ToString(),
                    patientId = dr["PATIENT_ID"].ToString(),
                    patName = dr["PAT_NAME"].ToString(),
                    patAge = dr["PAT_AGE"].ToString(),
                    surgeon = dr["SURGEON"].ToString(),
                    surgeonName = dr["SURGEON_NAME"].ToString(),
                    anesthesiaDoctor = dr["ANESTHESIA_DOCTOR"].ToString(),
                    anesthesiaDoctorName = dr["ANESTHESIA_DOCTOR_NAME"].ToString(),
                    anesthesiaMethod = dr["ANESTHESIA_METHOD"].ToString(),
                    operation = dr["OPERATION"].ToString(),
                    operationScale = dr["OPERATION_SCALE"].ToString(),
                    scheduleType = dr["SCHEDULE_TYPE"].ToString(),
                    notesOnOperation = dr["NOTES_ON_OPERATION"].ToString(),
                    firstAssistant = dr["FIRST_ASSISTANT"].ToString(),
                    firstAssistantName = dr["FIRST_ASSISTANT_NAME"].ToString(),
                    bedNo = dr["BED_NO"].ToString(),
                    operatingRoom = dr["OPERATING_ROOM"].ToString(),
                    sex = dr["SEX"].ToString(),
                    state = dr["STATE"].ToString().Trim()
                };
                if (info.state != "3")
                {
                    operatInfos.Add(info);
                }

            }

            if (operatInfos.Count == 0)
            {
                MessageBox.Show(datetime2+"没有手术排班记录!");
                return;
            }

          

            //需要推送给的主刀医生，一助，麻醉师
            List<string> userids = new List<string>();
            foreach (OperatInfo operatInfo in operatInfos)
            {
                if (!userids.Contains(operatInfo.surgeon))
                {
                    if (operatInfo.surgeon != null && !string.IsNullOrEmpty(operatInfo.surgeon))
                    {
                        userids.Add(operatInfo.surgeon);
                    }
                }
                if (!userids.Contains(operatInfo.firstAssistant))
                {
                    if (operatInfo.firstAssistant != null && !string.IsNullOrEmpty(operatInfo.firstAssistant))
                    {
                        userids.Add(operatInfo.firstAssistant);
                    }
                }
                if (!userids.Contains(operatInfo.anesthesiaDoctor))
                {
                    if (operatInfo.anesthesiaDoctor != null && !string.IsNullOrEmpty(operatInfo.anesthesiaDoctor))
                    {
                        userids.Add(operatInfo.anesthesiaDoctor);
                    }
                }
            }

            //推送给所有人企业微信通知
           // DateTime t1 = DateTime.Now;
            foreach (string userid in userids)
            {
                List<OperatInfo> infos = operatInfos.FindAll(x => x.surgeon == userid || x.firstAssistant == userid || x.anesthesiaDoctor == userid).OrderBy(r => r.sstime).ToList();

                string sendtext = string.Empty;
                sendtext += $"`您的手术排班通知`\r\n";
                sendtext += $"**事项详情:**  \r\n";
                sendtext += $"<font color=\"info\"> \r\n";
                int count = 1;
                foreach (OperatInfo info in infos)
                {
                    if (info.state == "0")
                    {
                        info.operatingRoomNoName = "未分配";
                        info.operation = "待确认";
                    }
                    sendtext += $"【" + (count++) + "】手术间号:" + info.operatingRoomNoName + "|第" + info.sequence + "台|" + info.patName + "、" + info.sex + "、" + info.patAge + "岁|" + info.operation + "|" + info.anesthesiaMethod + "|" + info.surgeonName; //+ "|" + info.sstime + "(手术时间)";
                    sendtext += "\r\n";
                }

                sendtext += $"</font>     \r\n";
                sendtext += $" \r\n";
                sendtext += $"日期：<font color=\"warning\">{datetime2}</font>  \n";
                //sendtext += $"时间：<font color=\"warning\">{infos[0].sstime}</font>  \r\n";
                sendtext += $" \r\n";
                sendtext += $"请合理安排时间准时参加手术!";
                sendtext += $" \r\n";

                string token = GetAccessToken();
                string msg = string.Empty;
                if (istest)
                {
                    //测试
                     msg= SendTextMessageMarkdown(sendUrl, token, textBox1.Text.ToString().Trim(), sendtext);
                }
                else
                {
                    //正式
                     msg = SendTextMessageMarkdown(sendUrl, token, userid, sendtext);
                }

                dynamic data = JsonConvert.DeserializeObject<dynamic>(msg);
                if (data.errmsg == "ok")
                {
                    foreach (OperatInfo operatInfo in infos)
                    {
                        operatInfo.是否推送成功 = "成功";
                        operatInfo.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                }
                else
                {
                    foreach (OperatInfo operatInfo in infos)
                    {
                        operatInfo.是否推送成功 = "失败";
                        operatInfo.推送时间 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                }
                string xxx=string.Empty;
                if (istest)
                {
                    xxx = textBox1.Text.ToString().Trim()+" (测试数据)" ;
                }
                else
                {
                    xxx = userid;
                }

                WriteLog("url:" + sendUrl + "    \r\n  token:" + token + "  \r\n userid:" + xxx + " \r\n  sendmessage:" + sendtext + " \r\n  回参:" + msg);
            }


            dataGridView1.DataSource = operatInfos;

            
        }







        //last 重复数据
        private void SendlastdayMessage(bool istest)
        {

            DateTime dateTime = DateTime.Now;
        
            dateTime = dateTime.AddDays(-1);
            string datetime = dateTime.ToString("yyyy-MM-dd");
            string datetime2 = dateTime.ToString("yyyy年MM月dd日");


            string sql = "SELECT A.*, B.NAME AS PAT_NAME, YEAR(getdate()) - YEAR(B.DATE_OF_BIRTH) AS PAT_AGE, B.INP_NO, B.SEX,ISNULL(rtrim(L.USER_NAME), " +
                "A.FIRST_ASSISTANT) AS FIRST_ASSISTANT_NAME, ISNULL(rtrim(M.USER_NAME), A.SECOND_ASSISTANT) AS SECOND_ASSISTANT_NAME, ISNULL(rtrim(N.USER_NAME), A.THIRD_ASSISTANT) AS THIRD_ASSISTANT_NAME," +
                "ISNULL(D.USER_NAME, A.SURGEON) AS SURGEON_NAME,  E.USER_NAME AS ANESTHESIA_DOCTOR_NAME, F.USER_NAME AS ANESTHESIA_ASSISTANT_NAME,  O.USER_NAME AS ANESTHESIA_ASSISTANT_NAME_TWO," +
                " G.USER_NAME AS FIRST_OPERATION_NURSE_NAME, H.USER_NAME AS SECOND_OPERATION_NURSE_NAME, I.USER_NAME AS FIRST_SUPPLY_NURSE_NAME, J.USER_NAME AS SECOND_SUPPLY_NURSE_NAME,  " +
                " K.USER_NAME AS THIRD_SUPPLY_NURSE_NAME, Q.DEPT_NAME AS WARD_CODE_NAME, case when A.VISIT_ID = 0 then '门诊' else '住院' end as SCHEDULE_TYPE ,   " +
                "    CASE WHEN ISNULL(RTRIM(OPERATING_ROOM_NO), '-1') = '-1' THEN  T.DEPT_NAME  ELSE  T.DEPT_NAME + '(' + CONVERT(VARCHAR, A.SEQUENCE) + ')' END AS OPERATING_DEPT_NAME,  CASE WHEN A.EMERGENCY_INDICATOR = 0 AND" +
                "  A.SCHEDULED_DATE_TIME IS NOT NULL AND              CONVERT(VARCHAR(10), P.ADMISSION_DATE_TIME, 121) =  CONVERT(VARCHAR(10), A.SCHEDULED_DATE_TIME, 121) THEN  1  ELSE   0  END AS DAYTIME_OPERATION,SS.*,MM.BED_LABEL FROM [155.155.100.107].docare.dbo.MED_OPERATION_SCHEDULE A" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_PAT_MASTER_INDEX B    ON A.PATIENT_ID = B.PATIENT_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS D   ON A.SURGEON = D.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS E    ON A.ANESTHESIA_DOCTOR = E.USER_ID  LEFT OUTER JOIN" +
                " [155.155.100.107].docare.dbo.MED_HIS_USERS F    ON A.ANESTHESIA_ASSISTANT = F.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS G    ON A.FIRST_OPERATION_NURSE = G.USER_ID  LEFT OUTER JOIN" +
                " [155.155.100.107].docare.dbo.MED_HIS_USERS H    ON A.SECOND_OPERATION_NURSE = H.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS I   ON A.FIRST_SUPPLY_NURSE = I.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS J  ON A.SECOND_SUPPLY_NURSE = J.USER_ID  LEFT " +
                "OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS K    ON A.THIRD_SUPPLY_NURSE = K.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS L    ON A.FIRST_ASSISTANT = L.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS M    ON A.SECOND_ASSISTANT = M.USER_ID" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS N    ON A.THIRD_ASSISTANT = N.USER_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_DEPT_DICT T    ON A.OPERATING_DEPT = T.DEPT_CODE  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_HIS_USERS O    ON A.SECOND_ANESTHESIA_ASSISTANT = O.USER_ID" +
                "  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_PATS_IN_HOSPITAL P    ON A.PATIENT_ID = P.PATIENT_ID   AND A.VISIT_ID = P.VISIT_ID  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_DEPT_DICT Q    ON P.WARD_CODE = Q.DEPT_CODE  LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_SCHEDULED_OPERATION_NAME SS ON A.PATIENT_ID =SS.PATIENT_ID " +
                " LEFT OUTER JOIN [155.155.100.107].docare.dbo.MED_SCHEDULED_OPERATION_NAME TT on SS.PATIENT_ID =TT.PATIENT_ID and SS.RESERVED1< TT.RESERVED1  left outer join [155.155.100.107].docare.dbo.MED_OPERATING_ROOM MM on  CONVERT(NVARCHAR(16), MM.BED_ID)=isnull(A.OPERATING_ROOM_NO,'-80')  WHERE TT.PATIENT_ID is NULL and  (CONVERT(VARCHAR, A.SCHEDULED_DATE_TIME, 23) =  @SheduleDateTime )" +
                "   AND A.OPERATING_ROOM = @OperatingRoom   AND NOT A.OPERATING_ROOM_NO IS NULL   AND A.STATE IN (0,1,2,3) ORDER BY ISNULL((SELECT BED_ID   FROM [155.155.100.107].docare.dbo.MED_OPERATING_ROOM   WHERE A.OPERATING_ROOM_NO = ROOM_NO),   0),   A.SEQUENCE";




            //住院手术室
            SqlParameter[] parameters = {
            new SqlParameter("@OperatingRoom", "1032"),
            new SqlParameter("@SheduleDateTime", datetime)
            };

            //日间手术室
            SqlParameter[] parameters2 = {
            new SqlParameter("@OperatingRoom", "2020"),
            new SqlParameter("@SheduleDateTime", datetime)
            };
            DataTable dt = DbHelper.GetData(sql, CommandType.Text, parameters);

            DataTable dt2 = DbHelper.GetData(sql, CommandType.Text, parameters2);

            dt.Merge(dt2);

            if (dt.Rows.Count == 0)
            {
                //MessageBox.Show(datetime2 + "没有手术排班记录!");
                return;
            }

            List<OperatInfo> operatInfos = new List<OperatInfo>();
            foreach (DataRow dr in dt.Rows)
            {
                DateTime dtt1 = DateTime.Parse(dr["SCHEDULED_DATE_TIME"].ToString());
                string resulttime = dtt1.ToString("HH:mm");

                OperatInfo info = new OperatInfo()
                {
                    operatingRoomNo = dr["OPERATING_ROOM_NO"].ToString(),
                    operatingRoomNoName = dr["BED_LABEL"].ToString(),
                    sequence = dr["SEQUENCE"].ToString(),

                    sstime = resulttime,
                    wardCodeName = dr["WARD_CODE_NAME"].ToString(),
                    operatdeptName = dr["OPERATING_DEPT_NAME"].ToString(),
                    inpNo = dr["INP_NO"].ToString(),
                    patientId = dr["PATIENT_ID"].ToString(),
                    patName = dr["PAT_NAME"].ToString(),
                    patAge = dr["PAT_AGE"].ToString(),
                    surgeon = dr["SURGEON"].ToString(),
                    surgeonName = dr["SURGEON_NAME"].ToString(),
                    anesthesiaDoctor = dr["ANESTHESIA_DOCTOR"].ToString(),
                    anesthesiaDoctorName = dr["ANESTHESIA_DOCTOR_NAME"].ToString(),
                    anesthesiaMethod = dr["ANESTHESIA_METHOD"].ToString(),
                    operation = dr["OPERATION"].ToString(),
                    operationScale = dr["OPERATION_SCALE"].ToString(),
                    scheduleType = dr["SCHEDULE_TYPE"].ToString(),
                    notesOnOperation = dr["NOTES_ON_OPERATION"].ToString(),
                    firstAssistant = dr["FIRST_ASSISTANT"].ToString(),
                    firstAssistantName = dr["FIRST_ASSISTANT_NAME"].ToString(),
                    bedNo = dr["BED_NO"].ToString(),
                    operatingRoom = dr["OPERATING_ROOM"].ToString(),
                    sex = dr["SEX"].ToString(),
                    state = dr["STATE"].ToString().Trim()
                };
                
                    operatInfos.Add(info);
                

            }

            if (operatInfos.Count == 0)
            {
                //MessageBox.Show(datetime2 + "没有手术排班记录!");
                return;
            }

            string patidstr= "'" + string.Join("','", operatInfos.Select(x => x.patientId)) + "'";
            string inpnostr = "'" + string.Join("','", operatInfos.Select(x => x.inpNo)) + "'";
            string bednostr = "'" + string.Join("','", operatInfos.Select(x => x.bedNo)) + "'";

            string sqll = "SELECT A.BED_NO,A.PATIENT_ID,A.SCHEDULED_DATE_TIME,A.ANESTHESIA_METHOD,A.OPERATION_NAME,A.SURGEON  from [155.155.100.107].docare.dbo.MED_OPERATION_MASTER A WHERE  ( CONVERT(VARCHAR, SCHEDULED_DATE_TIME, 23) < '" + datetime + "'  )  \r\nAND OPER_STATUS NOT IN (-80)  AND A.PATIENT_ID IN ("+patidstr+")  AND A.BED_NO IN ("+bednostr+")  ORDER BY A.OPERATING_ROOM_NO ASC,A.SEQUENCE ASC";


            SqlParameter[] parameters3 = {
            new SqlParameter("@OperatingRoom", "1032"),
            new SqlParameter("@SheduleDateTime", datetime)
            };
            DataTable dt3 = DbHelper.GetData(sqll, CommandType.Text, parameters3);

            List<repeatoperation> operatInfos2 = new List<repeatoperation>();

            foreach (DataRow dr in dt3.Rows)
            {
                DateTime dtt1 = DateTime.Parse(dr["SCHEDULED_DATE_TIME"].ToString());
                string resulttime = dtt1.ToString("yyyy年MM月dd日");

                repeatoperation info = new repeatoperation()
                {
                    
                    //inpNo = dr["INP_NO"].ToString(),
                    patientId = dr["PATIENT_ID"].ToString(),
                    //patName = dr["PAT_NAME"].ToString(),
                    surgeon = dr["SURGEON"].ToString(),
                    anesthesiaMethod = dr["ANESTHESIA_METHOD"].ToString(),
                    operation = dr["OPERATION_NAME"].ToString(),
                    
                    bedNo = dr["BED_NO"].ToString(),
                    sstime = resulttime,

                };
                operatInfos2.Add(info);
            }




            //推送给企业微信通知 管理者
    
            //List<OperatInfo> infos = operatInfos.FindAll(x => x.surgeon == userid || x.firstAssistant == userid || x.anesthesiaDoctor == userid).OrderBy(r => r.sstime).ToList();




            string sendtext = string.Empty;
            sendtext += $"`住院患者重复手术告知`\r\n";
            sendtext += $"**事项详情:**  \r\n";
            //sendtext += $"<font color=\"info\"> \r\n";
            int count = 1;
            foreach (OperatInfo info in operatInfos)
            {

                if (operatInfos2.Exists(r => r.patientId == info.patientId && r.bedNo == info.bedNo))
                {


                    sendtext += $"<font color=\"info\">\r\n ";
                    sendtext += $"【" + (count++) + "】手术间号:" + info.operatingRoomNoName + "|第" + info.sequence + "台|" + info.patName + "、" + info.sex + "、" + info.patAge + "岁|" + info.operation + "|" + info.anesthesiaMethod + "|" + info.surgeonName+"|"+ datetime2;
                    sendtext += $"</font>  ";
     
                    sendtext += $"<font color=\"warning\">  ";
                    sendtext += $"\r\n";
                    sendtext += "患者住院期间的其他手术:";
                    sendtext += $"</font>";
                    sendtext += $"\r\n";
                    int count2  = 1;
                    foreach (repeatoperation obj in operatInfos2.FindAll(r => r.patientId == info.patientId && r.bedNo == info.bedNo).OrderByDescending(r=>r.sstime).ToList())
                    {
                        sendtext += $"<font color=\"info\">";
                        sendtext +=  " <"+(count2++) +"> "+obj.operation+"|"+obj.anesthesiaMethod+"|"+obj.sstime;
                        sendtext += $"</font> \r\n";
                        sendtext += $"\r\n";
                    }
                    sendtext += $"\r\n";
                }
            }

            
            sendtext += $" \r\n";
            sendtext += $" \r\n";
            sendtext += $"<font color=\"black\">\r\n 推送日期：</font>   <font color=\"warning\">{DateTime.Now.ToString("yyyy年MM月dd日")}</font>  \n";
            //sendtext += $"时间：<font color=\"warning\">{infos[0].sstime}</font>  \r\n";
            sendtext += $" \r\n";
            sendtext += $"请手术管理者知晓上述手术情况!";
            sendtext += $" \r\n";

            string token = GetAccessToken();
            string msg = string.Empty;

            if (istest)
            {
                msg = SendTextMessageMarkdown(sendUrl, token, textBox1.Text.ToString().Trim(), sendtext);
            }
            else
            {
                msg = SendTextMessageMarkdown(sendUrl, token, "2382|1799|3024", sendtext);//推给高洁 手术管理人员 和本人
            }


            string xxx = string.Empty;
            if (istest)
            {
                xxx = textBox1.Text.ToString().Trim() + " (测试数据)";
            }
            else
            {
                xxx = "2382|1799|3024";
            }


            WriterepeatLog("url:" + sendUrl + "    \r\n  token:" + token + "  \r\n userid:" + xxx + " \r\n  sendmessage:" + sendtext + " \r\n  回参:" + msg);
            


        }


















        private void button2_Click(object sender, EventArgs e)
        {
            //富文本
            //SendTextMessageMarkdown(sendUrl, GetAccessToken(), "3024", "`您的手术排班通知`\r\n **事项详情:**  < font color =\"info\"> \r\n 【1】 7号楼-2号 /第1台/张三，女，44岁/左股骨骨折固定术/麻醉方式/ 张医生  \r\n 【2】 7号楼-2号/第2台/李四，女，87岁/左股骨骨折固定术/麻醉方式/张医生 \r\n 【3】 1号楼-3号/第4台/王五，女，88岁/左股骨骨折固定术/麻醉方式/高医生 </font>     \r\n \r\n 日　期：<font color=\"warning\">2025年5月16日</font>  \n 时　间：<font color=\"warning\">上午9:00-11:00</font>  \r\n \r\n 请合理安排时间准时参加手术!");

            this.button1.Enabled = false;
            this.button2.Enabled = false;

            //第二天
            SendMessage(true,true);

            this.button1.Enabled = true;
            this.button2.Enabled = true;

        }

        /// <summary>
        /// 每日早上7:00 推送当日的手术排班消息到个人企业微信通知，每日下午15:00推送第二日的手术排班消息到个人企业微信通知
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now.ToString("HH:mm") == "15:00" || DateTime.Now.ToString("HH:mm") == "15:01")
            {
                if (isEnabled)
                {
                    if (checkBox1.Checked)//
                    {
                        //下午三点推送第二天的手术排班
                        SendMessage(true, false);
                        
                    }
                    isEnabled = false;
                }
            }
            else if (DateTime.Now.ToString("HH:mm") == "07:00" || DateTime.Now.ToString("HH:mm") == "07:01")
            {
                if (isEnabled)
                {
                    if (checkBox1.Checked)//
                    {
                        //上午7点推送当天的手术排班
                        SendMessage(false, false);
                        
                    }
                    if (checkBox2.Checked)
                    {
                        //推送给管理员 前一天的手术排班消息有做多次手术的患者 推送给管理员
                        SendlastdayMessage(false);
                    }
                    isEnabled = false;
                }
            }
            else if (DateTime.Now.ToString("HH:mm") == "15:30" || DateTime.Now.ToString("HH:mm") == "15:31" || DateTime.Now.ToString("HH:mm") == "15:32" || DateTime.Now.ToString("HH:mm") == "07:30" || DateTime.Now.ToString("HH:mm") == "07:31" || DateTime.Now.ToString("HH:mm") == "07:32")
            {
                if (!isEnabled)
                {
                    isEnabled = true;
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            isEnabled = true;
            this.timer1.Interval = 90 * 1000; // 90秒轮询一次

            timer1.Start();

            button3.Enabled = false;
            button4.Enabled = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.timer1?.Dispose();
            //this.timer2?.Dispose();
        }

        /// <summary>
        /// 开启定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            this.timer1?.Start();
            //this.timer2?.Start();
            button3.Enabled = false;
            button4.Enabled = true;
        }

        /// <summary>
        /// 停止定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            this.timer1?.Stop();
            //this.timer2?.Stop();
            button3.Enabled = true;
            button4.Enabled = false;

        }

        /// <summary>
        /// 打开日志文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            // 获取路径（推荐AppDomain方式）
            string path = AppDomain.CurrentDomain.BaseDirectory+"log";

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

        private void button6_Click(object sender, EventArgs e)
        {
            SendlastdayMessage(true);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
