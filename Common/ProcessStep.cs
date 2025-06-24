using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace SendOperationPlan.Common
{
    public class ProcessStep
    {
        public string StepName { get; set; }
        public StepStatus Status { get; set; } = StepStatus.Pending;
    }

    public enum StepStatus
    {
        Pending,    // 待执行
        Running,    // 执行中
        Completed   // 已完成
    }

    public class ProgressReport
    {
        public List<ProcessStep> Steps { get; set; }
        public string CurrentStep { get; set; }
        public int CompletedCount { get; set; }
        public int PendingCount { get; set; }
    }






}
