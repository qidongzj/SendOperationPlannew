using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SendOperationPlan
{
    /// <summary>
    /// 危急值
    /// </summary>
    public class WjzInfo
    {
        public string 是否推送成功 { get; set; }

        //推送时间
        public string 推送时间 { get; set; }

        /// <summary>
        /// 1门诊  ， 2住院 
        /// </summary>
        public int CriticalType { get; set; }
        /// <summary>
        /// 患者姓名
        /// </summary>
        public string Hzxm { get; set; }
        
        /// <summary>
        /// 报告单号
        /// </summary>
        public string Bgdh { get; set; }

        /// <summary>
        /// 报告时间
        /// </summary>
        public string Bgsj { get; set; }

        /// <summary>
        /// 危急值序号  --门诊或住院危机值表的主序号 XH
        /// </summary>
        public decimal CriticalValue { get; set; }

        /// <summary>
        /// 插入时间
        /// </summary>
        public string SendTime { get; set; }

        /// <summary>
        /// 危急值内容
        /// </summary>
        public string Wjnr { get; set; } 

        /// <summary>
        /// 住院号码 住院用
        /// </summary>
        public string Zyhm { get; set; }   

        /// <summary>
        /// patid 门诊用
        /// </summary>
        public decimal Patid { get; set; }

        /// <summary>
        /// 送检开单医生代码
        /// </summary>
        public string Sjkdysdm { get; set; }
        /// <summary>
        /// 送检开单医生姓名
        /// </summary>
        public string Sjkdysmc { get; set; }

        /// <summary>
        /// 送检科室代码
        /// </summary>
        public string Sjksdm { get; set; }

        /// <summary>
        /// 送检科室名称
        /// </summary>
        public string Sjksmc { get; set; }

        /// <summary>
        /// 送检病区代码
        /// </summary>
        public string Sjbqdm { get; set; }

        /// <summary>
        /// 送检病区名称
        /// </summary>
        public string Sjbqmc { get; set; }

        /// <summary>
        /// 送检日期
        /// </summary>
        public string Sjrq { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 医生对危急值的处理意见
        /// </summary>
        public string Ysdfnr { get; set; }

        /// <summary>
        /// 医生答复状态  0未答复 1已答复
        /// </summary>
        public string Ysdfzt  { get; set; } 

        /// <summary>
        /// 接收医生代码
        /// </summary>
        public string Jsysdm { get; set; }

        /// <summary>
        /// 接收医生姓名
        /// </summary>
        public string Jsysmc { get; set; }

        /// <summary>
        /// 医生接收时间
        /// </summary>
        public string Ysjssj  { get; set; }


        /// <summary>
        /// 答复医生代码
        /// </summary>
        public string Dfysdm { get; set; }
        /// <summary>
        /// 答复医生名称
        /// </summary>
        public string Dfysmc { get; set; }
        /// <summary>
        /// 处置时间
        /// </summary>
        public string Czsj { get; set; } 

        /// <summary>
        /// 更新类型 1新增 2修改 
        /// </summary>
        public int Gxlx { get; set; }

       

    }
}
