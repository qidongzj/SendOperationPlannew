using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        public FormWjz()
        {
            InitializeComponent();
        }
    }
}
