using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SendOperationPlan.Common
{
    public partial class ProgressDialog : Form
    {
        public ProgressDialog()
        {
            InitializeComponent();
            InitializeComponent2();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void InitializeComponent2()
        {
            // 创建控件
            lblCurrent = new Label { AutoSize = true,ForeColor=  System.Drawing.Color.Red, Location = new Point(20, 20) , Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134))) };
            lblCompleted = new Label { AutoSize = true, Location = new Point(20, 50), Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134))) };
            lblPending = new Label { AutoSize = true, Location = new Point(20, 80), Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134))) };

            listSteps = new ListBox
            {
                ItemHeight = 22,
                Location = new Point(20, 110),
                Size = new Size(400, 350),
                Font= new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134))),
                DrawMode = DrawMode.OwnerDrawFixed
            };

            listSteps.DrawItem += (s, e) => {
                if (e.Index < 0) return;

                var step = (ProcessStep)listSteps.Items[e.Index];
                e.DrawBackground();
                System.Drawing.Brush brush;
                // 根据状态设置不同颜色
                switch (step.Status)
                {
                    case StepStatus.Completed:
                        brush = Brushes.Green;
                        break;
                    case StepStatus.Running:
                        brush = Brushes.Red;
                        break;
                    default:
                        brush = Brushes.Black;
                        break;
                }

                e.Graphics.DrawString(step.StepName, e.Font, brush, e.Bounds);
            };

            btnCancel = new Button
            {
                 Visible= false,//取消按钮是否显示
                Text = "取消",
                Location = new Point(440, 320),
                Size = new Size(100, 30),
                Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)))
            };

            btnCancel.Click += (s, e) => {
                if (MessageBox.Show("确定要中断操作吗？", "确认",
                    MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.DialogResult = DialogResult.Cancel;
                }
            };

            // 窗体属性
            this.Text = "业务流程执行状态";
            this.ClientSize = new Size(600, 480);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // 添加控件
            this.Controls.Add(lblCurrent);
            this.Controls.Add(lblCompleted);
            this.Controls.Add(lblPending);
            this.Controls.Add(listSteps);
            this.Controls.Add(btnCancel);
        }


        // 更新进度显示
        public void UpdateProgress(ProgressReport report)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<ProgressReport>(UpdateProgress), report);
                return;
            }

            // 更新当前步骤
            lblCurrent.Text = $"当前执行: {report.CurrentStep}";

            // 更新计数
            lblCompleted.Text = $"已完成: {report.CompletedCount}/{report.Steps.Count}";
            lblPending.Text = $"待执行: {report.PendingCount}";

            // 更新步骤列表
            listSteps.Items.Clear();
            foreach (var step in report.Steps)
            {
                listSteps.Items.Add(step);
            }
        }

        private ListBox listSteps;
        private Label lblCurrent;
        private Label lblCompleted;
        private Label lblPending;
        private Button btnCancel;
    }
}
