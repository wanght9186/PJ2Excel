using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using System.Net.NetworkInformation;
using Microsoft.Office.Core;
using System.Collections;
using System.Diagnostics;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using Mspj = Microsoft.Office.Interop.MSProject;
using Msxl = Microsoft.Office.Interop.Excel;
using System.Data.SqlTypes;

namespace Pj2Excel
{
    public static class TransToExcelFuc
    {
        public static void P2J_TransSelectionToExcel()
        {
            string currentMember = "";
            string lastMember = "";
            //获取选中的MsProject中的tasks
            Mspj.Tasks activeTasks = Utils.GetActiveTasks();

            List<MemberTaskInfo> mTaskInfoList = new List<MemberTaskInfo>();

            foreach (Mspj.Task tsk in activeTasks)
            {
                currentMember = tsk.ResourceNames.ToString();

                if (string.IsNullOrEmpty(currentMember))
                {
                    continue;
                }

                // 查找当前成员的任务信息，如果没有则创建新的
                MemberTaskInfo mTaskInfo = mTaskInfoList.FirstOrDefault(m => m.memberName == currentMember);
                if (mTaskInfo == null)
                {
                    mTaskInfo = new MemberTaskInfo
                    {
                        memberName = currentMember
                    };
                    mTaskInfoList.Add(mTaskInfo);
                }

                // 创建一个新的任务并添加到成员的任务列表中
                Task mTask;
                if (tsk.Text3 != null)
                {
                    mTask = new Task("【" + tsk.Text3 + "】 " + tsk.Name);
                }
                else
                {
                    mTask = new Task(tsk.Name);
                }

                // 判断是否有任务为 0 天，没有分段的情况
                if (tsk.SplitParts.Count != 0)
                {
                    // 把每段时间添加到任务中
                    foreach (Mspj.SplitPart part in tsk.SplitParts)
                    {
                        var start = part.Start;
                        var finish = part.Finish;
                        mTask.taskDurationList.Add(new TaskDuration(start, finish));
                    }
                }
                else
                {
                    var start = tsk.Start;
                    var finish = tsk.Finish;
                    mTask.taskDurationList.Add(new TaskDuration(start, finish));
                }

                mTaskInfo.taskList.Add(mTask);
            }

            string json = JsonConvert.SerializeObject(mTaskInfoList, Formatting.Indented);

            // 将 JSON 写入当前文件夹下的文件
            // File.WriteAllText("E:\\MemberExcelTask.json", json);

            TransToExcel(mTaskInfoList);

            Console.WriteLine("MemberExcelTask 对象已成功序列化并写入到 MemberExcelTask.json 文件。");
        }
        public static void TransToExcel(List<MemberTaskInfo> memberTaskList)
        {
            // 反序列化 JSON 数据
            List<MemberTaskInfo> memberExcelTaskList = memberTaskList;

            // 创建 Excel 文件
            Msxl.Application excelApp = new Msxl.Application();
            Msxl.Workbook workbook = excelApp.Workbooks.Add();
            Msxl.Worksheet worksheet = workbook.Sheets[1];

            // 设置行列宽度
            //Msxl.Range column = worksheet.Columns[1];
            //Msxl.Range row = worksheet.Rows[1];
            //column.ColumnWidth = 50;
            //row.RowHeight = 20;

            // 获取所有任务中的开始和结束日期
            DateTime minDate;
            DateTime maxDate;
            try
            {
                minDate = memberExcelTaskList
                    .SelectMany(member => member.taskList.SelectMany(task => task.taskDurationList.Select(td => (DateTime)td.start)))
                    .Min();
            }
            catch (Exception e)
            {
                minDate = DateTime.MinValue;
                MessageBox.Show("未找到【开始时间】");
            }

            try
            {
                maxDate = memberExcelTaskList
                    .SelectMany(member => member.taskList.SelectMany(task => task.taskDurationList.Select(td => (DateTime)td.finish)))
                    .Max();
            }
            catch (Exception)
            {
                maxDate = DateTime.MinValue;
                MessageBox.Show("未找到【结束时间】");
            }

            // 写入表头，日期从第二列开始
            int currentColumn = 2;
            for (DateTime currentDate = minDate; currentDate.Date <= maxDate.Date; currentDate = currentDate.AddDays(1))
            {
                worksheet.Cells[1, currentColumn].Value = currentDate.ToString("yyyy-MM-dd");
                worksheet.Columns[currentColumn].ColumnWidth = 50;
                currentColumn++;
            }

            // 写入每个 Name 的数据
            int currentRow = 2;
            foreach (var memberExcelTask in memberExcelTaskList)
            {
                foreach (var task in memberExcelTask.taskList)
                {
                    try
                    {
                        worksheet.Cells[currentRow, 1].Value = memberExcelTask.memberName;
                        int columnNum = 0;
                        foreach (var taskDuration in task.taskDurationList)
                        {
                            DateTime startDate = (DateTime)taskDuration.start;
                            DateTime endDate = (DateTime)taskDuration.finish;

                            for (DateTime currentDate = startDate; currentDate.Date <= endDate.Date; currentDate = currentDate.AddDays(1))
                            {
                                columnNum = GetColumnNumber(worksheet, currentDate, minDate);
                                if (!Utils.IsWeekend(currentDate.Date))
                                {
                                    worksheet.Cells[currentRow, columnNum].Value += "\n" + task.taskName;
                                }
                            }
                        }
                        //该任务的最后一天，单元格变成淡蓝色
                        Msxl.Range rng = worksheet.Cells[currentRow, columnNum];
                        rng.Value += "（---该条完结---）\n";
                        rng.Font.Color = Color.Red;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(task.taskName + "有问题");
                    }

                }


                worksheet.Columns[currentRow].RowHeight = 100;
                currentRow++;
            }

            // 保存 Excel 文件
            string excelFilePath = Path.Combine(Utils.GetPath(), $"时间导出_{minDate.ToString("yyyyMMdd")}-{maxDate.ToString("yyyyMMdd")}_{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx");
            workbook.SaveAs(excelFilePath);
            excelApp.Quit();
            MessageBox.Show("生成Daily完毕");
            Utils.OpenFolder(Utils.GetPath());

            Console.WriteLine($"Excel 文件已成功生成：{excelFilePath}");

        }

        // 获取日期对应的列号
        public static int GetColumnNumber(Msxl.Worksheet worksheet, DateTime date, DateTime minDate)
        {
            int startColumn = 2; // 日期从第二列开始
            int column = (int)(date.Date - minDate.Date).TotalDays + startColumn;
            return column;
        }
    }

    public class MemberTaskInfo
    {
        public string memberName;
        public List<Task> taskList = new List<Task>();
    }
    public class Task
    {
        public string taskName;
        public List<TaskDuration> taskDurationList;

        public Task()
        {
            this.taskName = "";
            taskDurationList = new List<TaskDuration>();
        }
        public Task(string taskName)
        {
            this.taskName = taskName;
            taskDurationList = new List<TaskDuration>();
        }
    }
    public class TaskDuration
    {
        public object start;
        public object finish;

        public TaskDuration(object start, object finish)
        {
            this.start = start;
            this.finish = finish;
        }
    }
}
