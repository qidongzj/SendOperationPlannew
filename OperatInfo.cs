using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendOperationPlan
{
    public class OperatInfo
    {
        public string 是否推送成功 { get; set; } 

        //推送时间
        public string 推送时间  { get; set; } 

        //手术间号
        public string operatingRoomNo { get; set; }
        //手术间号名称 BED_LABEL
        public string operatingRoomNoName { get; set; }
        //台次
        public string sequence { get; set; }
        //手术时间
        public string sstime { get; set; }

        public string sstime2 { get; set; } //手术时间2

        //病区代码
        public string wardCode { get; set; }
        //病区
        public string wardCodeName { get; set; }
        //科室
        public string operatdeptName { get; set; }

        public string operatdept { get; set; }
        //住院号
        public string inpNo { get; set; }
        //病人ID
        public string patientId { get; set; }
        //姓名
        public string patName { get; set; }
        //年龄
        public string patAge { get; set; }
        //性别
        public string sex { get; set; }
        //床号
        public string bedNo { get; set; }
        //手术名称
        public string operation { get; set; }
        //手术医师工号
        public string surgeon { get; set; }
        //手术医师
        public string surgeonName { get; set; }
        //麻醉方法
        public string anesthesiaMethod { get; set; }
        //手术预约类型
        public string scheduleType { get; set; }
        //备注
        public string notesOnOperation { get; set; }
        //一助工号
        public string firstAssistant { get; set; }
        //一助
        public string firstAssistantName { get; set; }
        //麻醉师工号
        public string anesthesiaDoctor { get; set; }
        //麻醉师
        public string anesthesiaDoctorName { get; set; }
        //状态
        public string state { get; set; }
        //等级
        public string operationScale { get; set; }
        //手术室
        public string operatingRoom { get; set; }


    }



    public class repeatoperation
    {
        //床号
        public string bedNo { get; set; }

        //病人ID
        public string patientId { get; set; }

        //手术时间
        public string sstime { get; set; }

        //麻醉方法
        public string anesthesiaMethod { get; set; }

        //手术医师工号
        public string surgeon { get; set; }

        //手术名称
        public string operation { get; set; }




    }
}
