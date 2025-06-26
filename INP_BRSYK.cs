using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendOperationPlan
{
    public class INP_BRSYK
    {
        public DateTime? RYRQ { get; set; } // 入院日期

        public DateTime? RQRQ { get; set; } // 入区日期

        public DateTime? CQRQ { get; set; } // 出区日期

        public DateTime? CYRQ { get; set; } // 出院日期    

        public decimal CISXH { get; set; } 

        public string PATID { get; set; } // 病人ID

        public string BLH { get; set; } // 病历号

       

        public string HZXM { get; set; } // 患者姓名    

        public string XSNL { get; set; } // 年龄

        public string SEX { get; set; } // 性别

        public string SFZH { get; set; } // 身份证号

      
        public string YBMC { get; set; } // 医保名称    

        public string KSDM { get; set; } // 科室代码
        public string KSMC { get; set; } // 科室名称

        public string BQDM { get; set; } // 病区代码    
        public string BQMC { get; set; } // 病区名称

        public string YSDM { get; set; } // 医生代码

        public string YSXM { get; set; } // 医生名称

        public string RYZDMC { get; set; } // 入院诊断

        public string CYZDMC { get; set; } // 出院诊断            


    }




    public class Result 
    {
        public string HZXM { get; set; } // 患者姓名

        public string XSNL { get; set; } // 年龄

        public string SEX { get; set; } // 性别

        public string SFZH { get; set; } // 身份证号

        public string YBMC { get; set; } // 医保名称    

        public DateTime CYRQ { get; set; } // 出院日期   

        public string CYKSMC  { get; set; } //出院科室

        public string RYKSMC { get; set; } // 入院科室

        public DateTime RYRQ { get; set; } // 入院日期

        public string SCZYH { get; set; } // 上次住院号

        public string BCZYH { get; set; } // 本次住院号

        public string RYZDMC { get; set; } // 入院诊断

        public string CYZDMC { get; set; } // 出院诊断

        public int ZYTS { get; set; } // 住院天数
    } 
}
