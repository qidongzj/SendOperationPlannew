using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using SendOperationPlan.Common;

namespace SendOperationPlan
{
    /// <summary>
    /// 中医优势病种查询
    /// </summary>
    public partial class FormZybz : Form
    {
        public FormZybz()
        {
            InitializeComponent();
            this.button2.Click += this.button2_Click;
            this.button3.Click += this.button3_Click;
            //checkedListBox1.DrawMode = DrawMode.OwnerDrawFixed;
            //checkedListBox1.DrawItem += CheckedListBox1_DrawItem;
            //checkedListBox1.ItemCheck += checkedListBox1_ItemCheck;
        }

        private List<zyInfo> zylist = new List<zyInfo>();

        private List<DataTable> dtlist = new List<DataTable>();



        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                // 选中时，设置为只读
                this.label1.Enabled = false;
                this.label2.Enabled = false;
                this.dateTimePicker1.Enabled = false;
                this.dateTimePicker2.Enabled = false;
                this.checkBox1.CheckedChanged -= checkBox1_CheckedChanged; // 移除事件处理器
                this.checkBox1.Checked = false; // 取消选中
                this.checkBox1.CheckedChanged += checkBox1_CheckedChanged; // 重新添加事件处理器
            }
            else
            {
                this.label1.Enabled = true;
                this.label2.Enabled = true;
                this.dateTimePicker1.Enabled = true;
                this.dateTimePicker2.Enabled = true;
                this.checkBox1.CheckedChanged -= checkBox1_CheckedChanged; // 移除事件处理器
                this.checkBox1.Checked = true; // 取消选中
                this.checkBox1.CheckedChanged += checkBox1_CheckedChanged; // 重新添加事件处理器

            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                this.label1.Enabled = true;
                this.label2.Enabled = true;
                this.dateTimePicker1.Enabled = true;
                this.dateTimePicker2.Enabled = true;
                this.checkBox2.CheckedChanged -= checkBox2_CheckedChanged; // 移除事件处理器
                this.checkBox2.Checked = false; // 取消选中
                this.checkBox2.CheckedChanged += checkBox2_CheckedChanged; // 重新添加事件处理器

            }
            else
            {
                this.label1.Enabled = false;
                this.label2.Enabled = false;
                this.dateTimePicker1.Enabled = false;
                this.dateTimePicker2.Enabled = false;
                this.checkBox2.CheckedChanged -= checkBox2_CheckedChanged; // 移除事件处理器
                this.checkBox2.Checked = true; // 取消选中
                this.checkBox2.CheckedChanged += checkBox2_CheckedChanged; // 重新添加事件处理器
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.button1.Text == "全选")
            {
                // 全选
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, true); // 勾选第 i 项



                }
                this.button1.Text = "取消全选";
            }
            else
            {
                // 取消全选
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false); // 勾选第 i 项
                }
                this.button1.Text = "全选";
            }
        }

        private void FormZybz_Load(object sender, EventArgs e)
        {
            //checkedListBox1.ItemHeight = 30; 
            this.checkBox2.Checked = true;
            zylist.Add(new zyInfo() { mbdm = "EMR.5739978933107995309", mbmc = "肛痈" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5758688258368588003", mbmc = "混合痔" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5065277266013166612", mbmc = "心水" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5664678132915010726", mbmc = "肝癖" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5750893277970205727", mbmc = "休息痢" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5346920438105085672", mbmc = "腹痛" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5688732593386452743", mbmc = "泄泻" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4658996579313775027", mbmc = "腰痹" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5386519928306777482", mbmc = "膝痹" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5495072453092167063", mbmc = "颈椎病" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5667975912105643807", mbmc = "漏肩风" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4837268271268056888", mbmc = "桡骨" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5197837065423282692", mbmc = "锁骨" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5276576165693924800", mbmc = "体格检查" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4647765519865822094", mbmc = "慢性肾病" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5407296338640325339", mbmc = "消渴" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4881269683808916808", mbmc = "臁疮" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5022869359371481608", mbmc = "下消" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5288136671363637513", mbmc = "盆腔炎" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5362562035230070307", mbmc = "胎动不安" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5731057216148498847", mbmc = "眩晕" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4946257442637255693", mbmc = "风温" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4987122773844770477", mbmc = "丹毒" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5615494466810242448", mbmc = "口僻" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5374245403166413436", mbmc = "肺瘅" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5018082970636394437", mbmc = "内伤咳嗽" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4635300756876562918", mbmc = "肺癌" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5681420302526664162", mbmc = "尪痹" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5686306738709052388", mbmc = "尪痹HAQ" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5693797141667994497", mbmc = "蛇串疮" });
            zylist.Add(new zyInfo() { mbdm = "EMR.4820497941642217375", mbmc = "脱疽" });
            zylist.Add(new zyInfo() { mbdm = "EMR.5343596212353785445", mbmc = "热淋" });

            zylist.Add(new zyInfo() { mbdm = "EMR.5717803281003616614", mbmc = "劳淋" });

            zylist.Add(new zyInfo() { mbdm = "EMR.5035569906529502489", mbmc = "胃癌" });

        }
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            checkedListBox1.Invalidate(); // 强制重绘
            // 仅重绘当前变更项
            //Rectangle itemRect = checkedListBox1.GetItemRectangle(e.Index);
            //checkedListBox1.Invalidate(itemRect);

        }

        private void CheckedListBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            // 获取当前项的选中状态
            bool isChecked = checkedListBox1.GetItemChecked(e.Index);
            // 选中项用红色，未选中项用默认颜色
            Brush textBrush = isChecked ? Brushes.Red : new SolidBrush(e.ForeColor);

            // 绘制背景（避免选中背景覆盖）
            e.DrawBackground();

            // 绘制文本
            e.Graphics.DrawString(
                checkedListBox1.Items[e.Index].ToString(),
                e.Font,
                textBrush,
                e.Bounds,
                StringFormat.GenericDefault
            );

            // 绘制焦点框
            e.DrawFocusRectangle();
        }




        //public  void RunTasksWithProgress(List<string> steps, Action<string> processStep)
        //{
        //    var progressForm = new ProgressForm(steps);
        //    var worker = new BackgroundWorker { WorkerSupportsCancellation = true };
        //    // 绑定事件  
        //    worker.DoWork += (sender, e) =>
        //    {
        //        for (int i = 0; i < steps.Count; i++)
        //        {
        //            if (worker.CancellationPending) { e.Cancel = true; return; }
        //            string step = steps[i];
        //            // 报告当前步骤  
        //            worker.ReportProgress(i, step);
        //            processStep(step); // 执行业务逻辑  
        //        }
        //    };
        //    // 更新 UI  
        //    worker.ProgressChanged += (sender, e) =>
        //    {
        //        progressForm.UpdateCurrentStep(e.ProgressPercentage, (string)e.UserState);
        //    };

        //    worker.RunWorkerCompleted += (sender, e) =>
        //    {
        //        progressForm.Close();
        //        if (e.Cancelled) MessageBox.Show("任务已取消");
        //    };



        //    // 取消回调  
        //    progressForm.CancelRequested = () => worker.CancelAsync();
        //    progressForm.ShowDialog(); // 显示进度窗体  

        //    worker.RunWorkerAsync(); // 启动任务 

        //}

        private async Task<List<DataTable>> ProcessStepExecuteasync(string mbmc,int yy)
        {
            if (yy == 1)
            {

                if (checkBox2.Checked)
                {

                    DataTable dt = await DbHelper.QueryDataAsync(" exec CISDB.dbo.GetZylxInfo '" + zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm + "'");
                    dtlist.Add(dt);
                    return dtlist;
                }
                else
                {
                    DataTable dt = await DbHelper.QueryDataAsync(" exec CISDB.dbo.GetZylxInfoForDate '" + zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm + "','" + dateTimePicker1.Value.ToString("yyyy-MM-dd hh:mm:ss") + "','" + dateTimePicker2.Value.ToString("yyyy-MM-dd hh:mm:ss") + "'");
                    dtlist.Add(dt);
                    return dtlist;

                }
            }
            else 
            {
                DataSet ds = null;
                if (checkBox2.Checked)
                {
                     ds = await DbHelper.QueryDataSetAsync(" exec CISDB.dbo.GetdifferentZyysbz '" + zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm + "'");
                }
                else 
                {
                     ds = await DbHelper.QueryDataSetAsync(" exec CISDB.dbo.GetdifferentZyysbzForDate '" + zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm + "','" + dateTimePicker1.Value.ToString("yyyy-MM-dd hh:mm:ss") + "','" + dateTimePicker2.Value.ToString("yyyy-MM-dd hh:mm:ss") + "'");
                }
                //GetdifferentZyysbzForDate
                if (ds.Tables.Count == 2)
                {
                    //foreach (DataTable dt in ds.Tables)
                    //{
                    //    dtlist.Add(dt);


                    //}
                    if (ds.Tables[0].Rows.Count == 0)
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("数据");
                        dt.Columns.Add("模板代码");
                        dt.Columns.Add("模板名称");
                        dt.Columns.Add("EMRXH");
                        DataRow dr = dt.NewRow();
                        dr["数据"] = "数据不存在";
                        dr["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                        dr["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                        dr["EMRXH"] = "-99"; // 添加EMRXH列
                        dt.Rows.Add(dr);
                        //ds.Tables[0] = dt;

                        //比较下空重取

                        DataTable dt_temp = FilterByLinq(dt, ds.Tables[1]);
                        if (dt_temp == null || dt_temp.Rows.Count == 0)
                        {
                            // dt_temp.Columns.Add("数据");
                            // dt_temp.Columns.Add("模板代码");
                            // dt_temp.Columns.Add("模板名称");
                            // dt_temp.Columns.Add("EMRXH");
                            DataRow dr1 = dt_temp.NewRow();
                            dr1["数据"] = "数据不存在";
                            dr1["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                            dr1["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                            dr1["EMRXH"] = "-99"; // 添加EMRXH列
                            dt_temp.Rows.Add(dr1);
                        }
                        dtlist.Add(dt_temp);
                    }
                    else
                    {
                        //比较下空重取
                        DataTable dt_temp = FilterByLinq(ds.Tables[0], ds.Tables[1]);
                        if (dt_temp == null || dt_temp.Rows.Count == 0)
                        {
                            dt_temp.Columns.Add("数据");
                            //dt_temp.Columns.Add("模板代码");
                            dt_temp.Columns.Add("模板名称");
                            //dt_temp.Columns.Add("EMRXH");
                            DataRow dr1 = dt_temp.NewRow();
                            dr1["数据"] = "数据不存在";
                            dr1["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                            dr1["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                            dr1["EMRXH"] = "-99"; // 添加EMRXH列
                            dt_temp.Rows.Add(dr1);
                        }

                        dtlist.Add(dt_temp);
                    }

                    if (ds.Tables[1].Rows.Count == 0)
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("数据");
                        dt.Columns.Add("模板代码");
                        dt.Columns.Add("模板名称");
                        dt.Columns.Add("EMRXH");
                        DataRow dr = dt.NewRow();
                        dr["数据"] = "数据不存在";
                        dr["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                        dr["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                        dr["EMRXH"] = "-33"; // 添加EMRXH列
                        dt.Rows.Add(dr);
                        //比较下空重取
                        //dtlist.Add(FilterByLinq(dt, ds.Tables[0]));

                        DataTable dt_temp = FilterByLinq(dt, ds.Tables[0]);
                        if (dt_temp == null || dt_temp.Rows.Count == 0)
                        {
                            //dt_temp.Columns.Add("数据");
                            //dt_temp.Columns.Add("模板代码");
                            //dt_temp.Columns.Add("模板名称");
                            //dt_temp.Columns.Add("EMRXH");
                            DataRow dr1 = dt_temp.NewRow();
                            dr1["数据"] = "数据不存在";
                            dr1["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                            dr1["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                            dr1["EMRXH"] = "-33"; // 添加EMRXH列
                            dt_temp.Rows.Add(dr1);
                        }

                        dtlist.Add(dt_temp);

                    }
                    else
                    {

                        DataTable dt_temp = FilterByLinq(ds.Tables[1], ds.Tables[0]);
                        if (dt_temp == null || dt_temp.Rows.Count == 0)
                        {
                            dt_temp.Columns.Add("数据");
                            //dt_temp.Columns.Add("模板代码");
                            //dt_temp.Columns.Add("模板名称");
                            //dt_temp.Columns.Add("EMRXH");
                            DataRow dr1 = dt_temp.NewRow();
                            dr1["数据"] = "数据不存在";
                            dr1["模板代码"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbdm;
                            dr1["模板名称"] = zylist.FirstOrDefault(r => r.mbmc == mbmc).mbmc;
                            dr1["EMRXH"] = "-33"; // 添加EMRXH列
                            dt_temp.Rows.Add(dr1);
                        }

                        dtlist.Add(dt_temp);
                        //比较下空重取
                        //dtlist.Add(FilterByLinq(ds.Tables[1], ds.Tables[0]));
                    }





                }

                return dtlist;
            }
        }




    private DataTable FilterByLinq(DataTable dt1, DataTable dt2)
    {
        // 获取 dt1 中所有 dt2 不存在的行（差集）
        var diffRows = dt1.AsEnumerable()
            .Where(row1 => !dt2.AsEnumerable()
                .Any(row2 => row2["EMRXH"].Equals(row1["EMRXH"])));

        // 转换为新 DataTable
        DataTable result = dt1.Clone();
        foreach (DataRow row in diffRows)
            result.ImportRow(row);

        return result;
    }




    //public async Task<DataTable> ProcessStepExecuteAsync()
    //{
    //    var users = new DataTable();

    //    using (var connection = new SqlConnection(ConnectionString))
    //    {
    //        await connection.OpenAsync(); // 异步打开连接[6,7](@ref)

    //        using (var command = new SqlCommand("SELECT Id, Name, Email FROM Users", connection))
    //        using (var reader = await command.ExecuteReaderAsync()) // 异步读取数据[6](@ref)
    //        {
    //            while (await reader.ReadAsync()) // 异步逐行读取[7](@ref)
    //            {
    //                users.Add(new User
    //                {
    //                    Id = reader.GetInt32(0),
    //                    Name = reader.GetString(1),
    //                    Email = reader.GetString(2)
    //                });
    //            }
    //        }
    //    }

    //    return users;
    //}


    public async Task ExecuteAsync(IProgress<ProgressReport> progress, List<ProcessStep> steps,int yy)
        {
            //var steps = new List<ProcessStep> {
            //new ProcessStep { StepName = "数据预处理" },
            //new ProcessStep { StepName = "执行计算" },
            //new ProcessStep { StepName = "结果验证" },
            //new ProcessStep { StepName = "数据保存" },
            //new ProcessStep { StepName = "生成报告" }};

            for (int i = 0; i < steps.Count; i++)
            {
                // 更新步骤状态
                steps[i].Status = StepStatus.Running;

                // 生成进度报告
                var report = new ProgressReport
                {
                    Steps = new List<ProcessStep>(steps),
                    CurrentStep = steps[i].StepName,
                    CompletedCount = steps.Count(s => s.Status == StepStatus.Completed),
                    PendingCount = steps.Count(s => s.Status == StepStatus.Pending)
                };

                progress?.Report(report);

                // 模拟业务执行（实际替换为业务代码）
                //  await Task.Delay(new Random().Next(1000, 3000));
                //ProcessStepExecute(report.CurrentStep);
                await ProcessStepExecuteasync(report.CurrentStep,yy);

                // 标记完成
                steps[i].Status = StepStatus.Completed;

                // 更新完成后的状态
                report.CompletedCount++;
                report.PendingCount = steps.Count - report.CompletedCount - 1;
                progress?.Report(report);
            }
        }


        private async Task StartProcessAsync(List<ProcessStep> mbmclist,int yy)
        {
            using (var progressDialog = new ProgressDialog())
            {
                // 创建进度报告器
                var progress = new Progress<ProgressReport>(report =>
                    progressDialog.UpdateProgress(report)
                );

                // 显示非阻塞弹窗
                progressDialog.Show(this);

                try
                {

                   

                    // 执行业务流程
                    await ExecuteAsync(progress, mbmclist,yy);

                    if (progressDialog.DialogResult != DialogResult.Cancel)
                    {
                        MessageBox.Show("所有业务流程已完成！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"执行出错: {ex.Message}");
                }
                finally
                {
                    progressDialog.Close();
                }
            }
        }



        private async void button2_Click(object sender, EventArgs e)
        {


            try
            {
                //if (checkBox2.Checked)
                //{
                //不限制时间

                if (checkedListBox1.CheckedItems.Count == 0)
                {
                    MessageBox.Show("请至少选择一个病种！");
                    return;
                }

                IsEnable(false);
                // this.button2.Enabled = false;
                // this.button3.Enabled = false;

                //for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                //{
                //    string mbdm = zylist.FirstOrDefault(r => r.mbmc == checkedListBox1.CheckedItems[i].ToString()).mbdm;
                //    DataTable dt =  await DbHelper.QueryDataAsync(" exec CISDB.dbo.GetZylxInfo '" + mbdm + "'");
                //    dtlist.Add(dt);
                //}


                List<ProcessStep> mbmclist = new List<ProcessStep>();

                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    string mbmc = zylist.FirstOrDefault(r => r.mbmc == checkedListBox1.CheckedItems[i].ToString()).mbmc;

                    mbmclist.Add(new ProcessStep() { StepName = mbmc, Status = StepStatus.Pending });
                }
                dtlist = new List<DataTable>();

                await StartProcessAsync(mbmclist,1);

                //RunTasksWithProgress(mbmclist, ProcessStepExecute);

                // ProgressForm fm = new ProgressForm();
                // fm.RunTasksWithProgress(mbmclist, ProcessStepExecute);

                if (dtlist.Count > 0)
                {
                    string sj = DateTime.Now.ToString("yyyyMMddHH:mm:ss").Trim().Replace(":", "");
                    // 导出到Excel
                    string ryrq = string.Empty;
                    if (checkBox2.Checked)
                    {
                        ryrq = string.Empty;
                    }
                    else
                    {
                        ryrq = "(" + dateTimePicker1.Value.ToString("yyyyMMdd") + "-" + dateTimePicker2.Value.ToString("yyyyMMdd") + ")";
                    }

                    ExportToExcel(dtlist, "D:\\中医优势病种数据查询集" + sj + ryrq + ".xlsx");
                    MessageBox.Show("导出成功！　D:\\中医优势病种数据查询集" + sj + ryrq + ".xlsx");

                    string msg = "D:\\中医优势病种数据查询集" + sj + ryrq + ".xlsx";

                    System.Diagnostics.Process.Start(msg); // 打开文件夹并选中导出的文件
                }

                IsEnable(true);

                //this.button2.Enabled = true;// 恢复按钮状态
                //this.button3.Enabled = true;
                //}
                //else
                //{
                //    this.button2.Enabled = false;



                //    this.button2.Enabled = true;// 恢复按钮状态

                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("异常：" + ex.Message);
                // this.button2.Enabled = true; // 恢复按钮状态
                // this.button3.Enabled = true;
                IsEnable(true);
            }


        }


        public void ExportToExcel(List<DataTable> dtlist, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                //var worksheet = workbook.Worksheets.Add(dt, "Sheet1");
                //worksheet.Columns().AdjustToContents(); // 自动调整列宽

                foreach (DataTable dt in dtlist)
                {
                    string mm = string.Empty;
                    if (dt.Rows[0][0].ToString() == "数据不存在")
                    {
                        mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc + "（无数据）";
                    }
                    else
                    {
                        mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc;
                        //if (dt.Rows[0]["模板名称"].ToString().Trim().Length > 31)
                        //{
                        //    mm = dt.Rows[0]["模板名称"].ToString().Trim().Substring(0, 31);
                        //}
                        //else
                        //{
                        //    mm = dt.Rows[0]["模板名称"].ToString().Trim();
                        //}

                    }
                    var worksheet = workbook.Worksheets.Add(dt, mm);
                    worksheet.Columns().AdjustToContents(); // 自动调整列宽
                                                            //worksheet.Rows().AddHorizontalPageBreaks();
                }
                workbook.SaveAs(filePath);

            }
        }




        public void ExportToExcel2(List<DataTable> dtlist, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                //var worksheet = workbook.Worksheets.Add(dt, "Sheet1");
                //worksheet.Columns().AdjustToContents(); // 自动调整列宽

                foreach (DataTable dt in dtlist)
                {
                    string mm = string.Empty;
                    //if (dt.Rows[0][0].ToString() == "数据不存在")
                    //{
                    //    mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc ;
                    //}
                    //else
                    //{
                    //    mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc;

                    //}

                    string sheetName = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc + "(诊断满足但未有对应优势病种病历的患者)";
                    bool sheetExists = workbook.Worksheets.Any(sheet =>sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                    if (!sheetExists)
                    {

                        //mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc+ "(诊断满足但未有对应优势病种病历的患者)";
                        var worksheet = workbook.Worksheets.Add(dt, sheetName);
                        worksheet.Columns().AdjustToContents(); // 自动调整列宽
                        
                    }
                    else 
                    {
                        //mm = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc + "(已有对应的优势病历但诊断条件不满足的)";

                        string sheetName1 = zylist.FirstOrDefault(r => r.mbdm == dt.Rows[0]["模板代码"].ToString().Trim()).mbmc + "(已有对应的优势病历但诊断条件不满足的)";
                        bool sheetExists1 = workbook.Worksheets.Any(sheet => sheet.Name.Equals(sheetName1, StringComparison.OrdinalIgnoreCase));

                        if (!sheetExists1)
                        {
                            var worksheet = workbook.Worksheets.Add(dt, sheetName1);
                            worksheet.Columns().AdjustToContents(); // 自动调整列宽
                        }
                    }
                }
                workbook.SaveAs(filePath);

            }
        }


        private void IsEnable(bool T)
        {
            this.button1.Enabled = T;
            this.button2.Enabled = T;
            this.button3.Enabled = T;
            this.checkBox1.Enabled = T;
            this.checkBox2.Enabled = T;
            this.dateTimePicker1.Enabled = T;
            this.dateTimePicker2.Enabled = T;
            this.checkedListBox1.Enabled = T;

        }


        private void checkedListBox1_MouseClick(object sender, MouseEventArgs e)
        {
            int index = checkedListBox1.IndexFromPoint(e.Location);
            if (index >= 0)
            {
                bool isChecked = !checkedListBox1.GetItemChecked(index);
                checkedListBox1.SetItemChecked(index, isChecked);

            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {

            try
            {
                if (checkedListBox1.CheckedItems.Count == 0)
                {
                    MessageBox.Show("请至少选择一个病种！");
                    return;
                }

                IsEnable(false);
                //this.button3.Enabled = false;
                //this.button2.Enabled = false; 

                List<ProcessStep> mbmclist = new List<ProcessStep>();

                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    string mbmc = zylist.FirstOrDefault(r => r.mbmc == checkedListBox1.CheckedItems[i].ToString()).mbmc;

                    mbmclist.Add(new ProcessStep() { StepName = mbmc, Status = StepStatus.Pending });
                }
                dtlist = new List<DataTable>();

                await StartProcessAsync(mbmclist,2);//校验

                if (dtlist.Count > 0)
                {
                    string sj = DateTime.Now.ToString("yyyyMMddHH:mm:ss").Trim().Replace(":", "");
                    // 导出到Excel
                    string ryrq = string.Empty;
                    if (checkBox2.Checked)
                    {
                        ryrq = string.Empty;
                    }
                    else
                    {
                        ryrq = "(" + dateTimePicker1.Value.ToString("yyyyMMdd") + "-" + dateTimePicker2.Value.ToString("yyyyMMdd") + ")";
                    }

                    ExportToExcel2(dtlist, "D:\\(校验)中医优势病种数据校验" + sj + ryrq + ".xlsx");
                    MessageBox.Show("导出成功！　D:\\(校验)中医优势病种数据校验" + sj + ryrq + ".xlsx");

                    string msg = "D:\\(校验)中医优势病种数据校验" + sj + ryrq + ".xlsx";

                    System.Diagnostics.Process.Start(msg); // 打开文件夹并选中导出的文件
                }

                IsEnable(true);
                //this.button3.Enabled = true;// 恢复按钮状态
                //this.button2.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("异常：" + ex.Message);
                IsEnable(true);
                //this.button3.Enabled = true; // 恢复按钮状态
                //this.button2.Enabled = true;
            }
        } 
    }


    
}
