namespace SendOperationPlan.Common
{
    partial class ProgressForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblCurrentTask = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBarTotal = new System.Windows.Forms.ProgressBar();
            this.progressBarItem = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.lstTaskStatus = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblCurrentTask
            // 
            this.lblCurrentTask.AutoSize = true;
            this.lblCurrentTask.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCurrentTask.Location = new System.Drawing.Point(272, 315);
            this.lblCurrentTask.Name = "lblCurrentTask";
            this.lblCurrentTask.Size = new System.Drawing.Size(149, 20);
            this.lblCurrentTask.TabIndex = 0;
            this.lblCurrentTask.Text = "lblCurrentTask";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(41, 315);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(179, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "当前正在执行任务:";
            // 
            // progressBarTotal
            // 
            this.progressBarTotal.Location = new System.Drawing.Point(197, 12);
            this.progressBarTotal.Name = "progressBarTotal";
            this.progressBarTotal.Size = new System.Drawing.Size(440, 23);
            this.progressBarTotal.TabIndex = 2;
            this.progressBarTotal.Visible = false;
            // 
            // progressBarItem
            // 
            this.progressBarItem.Location = new System.Drawing.Point(264, 351);
            this.progressBarItem.Name = "progressBarItem";
            this.progressBarItem.Size = new System.Drawing.Size(362, 23);
            this.progressBarItem.TabIndex = 3;
            this.progressBarItem.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(62, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "总任务进度：";
            this.label3.Visible = false;
            // 
            // lstTaskStatus
            // 
            this.lstTaskStatus.FormattingEnabled = true;
            this.lstTaskStatus.ItemHeight = 12;
            this.lstTaskStatus.Location = new System.Drawing.Point(55, 78);
            this.lstTaskStatus.Name = "lstTaskStatus";
            this.lstTaskStatus.Size = new System.Drawing.Size(571, 232);
            this.lstTaskStatus.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(39, 351);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(219, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "当前正在执行任务进度:";
            this.label1.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(51, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 20);
            this.label4.TabIndex = 7;
            this.label4.Text = "任务列表";
            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(640, 386);
            this.ControlBox = false;
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lstTaskStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBarItem);
            this.Controls.Add(this.progressBarTotal);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblCurrentTask);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.Text = "ProgressForm";
            this.Load += new System.EventHandler(this.ProgressForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblCurrentTask;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBarTotal;
        private System.Windows.Forms.ProgressBar progressBarItem;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstTaskStatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
    }
}