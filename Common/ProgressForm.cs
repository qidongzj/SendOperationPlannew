using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendOperationPlan.Common
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
        }
        public Action CancelRequested { get; set; }
        private readonly List<string> _steps;
        public ProgressForm(List<string> steps)
        {
            InitializeComponent();
            _steps = steps;
            lstTaskStatus.DataSource = steps; // 绑定步骤列表  
        }
        // 更新当前步骤  
        public void UpdateCurrentStep(int index, string stepName)
        {
            lstTaskStatus.SelectedIndex = index; // 高亮当前步骤  
            lblCurrentTask.Text = $"当前: {stepName}";
        }
        // 取消按钮事件  
        private void btnCancel_Click(object sender, EventArgs e) => CancelRequested?.Invoke();


        /*
        public void RunTasksWithProgress(List<string> steps, Action<string> processStep)
        {
            var progressForm = new ProgressForm(steps);
            var worker = new BackgroundWorker { WorkerSupportsCancellation = true };
            // 绑定事件  
            worker.DoWork += (sender, e) =>
            {
                for (int i = 0; i < steps.Count; i++)
                {
                    if (worker.CancellationPending) { e.Cancel = true; return; }
                    string step = steps[i];
                    // 报告当前步骤  
                    worker.ReportProgress(i, step);
                    processStep(step); // 执行业务逻辑  
                }
            };
            // 更新 UI  
            worker.ProgressChanged += (sender, e) =>
            {
                progressForm.UpdateCurrentStep(e.ProgressPercentage, (string)e.UserState);
            };

            worker.RunWorkerCompleted += (sender, e) =>
            {
                progressForm.Close();
                if (e.Cancelled) MessageBox.Show("任务已取消");
            };


            worker.RunWorkerAsync(); // 启动任务 
            // 取消回调  
            progressForm.CancelRequested = () => worker.CancelAsync();
            progressForm.ShowDialog(); // 显示进度窗体  

            

        }

        */


        private void ProgressForm_Load(object sender, EventArgs e)
        {
            //worker.RunWorkerAsync(); // 启动任务 
        }
    }

}
