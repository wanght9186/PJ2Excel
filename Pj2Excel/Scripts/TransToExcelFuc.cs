using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools;
using Mspj = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;
using System.Net.NetworkInformation;
using Microsoft.Office.Core;
//using Microsoft.Office.Interop.MSProject;
using Msxl = Microsoft.Office.Interop.Excel;
using System.Collections;
using Newtonsoft.Json;
using System.IO;
//using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;

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

                if (lastMember == "")
                {
                    lastMember = tsk.ResourceNames.ToString();
                }

                if (currentMember != lastMember)
                {
                    MemberTaskInfo mTaskInfo = new MemberTaskInfo();
                    mTaskInfoList.Add(mTaskInfo);
                    mTaskInfo.memberName = currentMember;

                    //一个task
                    Task mTask = new Task();
                    mTaskInfo.taskList.Add(mTask);
                    mTask.taskName = "【" + tsk.Text3 + "】 " + tsk.Name;

                    //判断是否有任务为O天，没有分段的情况
                    if (tsk.SplitParts.Count != 0)
                    {
                        //把每段时间添加到task中
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

                    lastMember = currentMember;
                }
                else
                {
                    Task mTask = new Task();
                    mTask.taskName = "【"+tsk.Text3 + "】 " + tsk.Name;
                    //判断是否有任务为O天，没有分段的情况
                    if (tsk.SplitParts.Count != 0)
                    {
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

                    if (mTaskInfoList.Count == 0)
                    {
                        //第一次
                        MemberTaskInfo mTaskInfo = new MemberTaskInfo();
                        mTaskInfo.memberName = currentMember;
                        mTaskInfoList.Add(mTaskInfo);
                        mTaskInfo.taskList.Add(mTask);
                    }
                    else
                    {
                        mTaskInfoList[mTaskInfoList.Count - 1].taskList.Add(mTask);
                    }
                }
            }

            string json = JsonConvert.SerializeObject(mTaskInfoList, Formatting.Indented);

            // 将 JSON 写入当前文件夹下的文件
            //File.WriteAllText("E:\\MemberExcelTask.json", json);

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
                                if (!IsWeekend(currentDate.Date))
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
            OpenFolder(Utils.GetPath());

            Console.WriteLine($"Excel 文件已成功生成：{excelFilePath}");

        }
        // 判断一个日期是否为周末
        public static bool IsWeekend(DateTime date)
        {
            // 判断日期的 DayOfWeek 属性是否为 Saturday 或 Sunday
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }
        // 获取日期对应的列号
        public static int GetColumnNumber(Msxl.Worksheet worksheet, DateTime date, DateTime minDate)
        {
            int startColumn = 2; // 日期从第二列开始
            int column = (int)(date.Date - minDate.Date).TotalDays + startColumn;
            return column;
        }
        //打开文件夹
        static void OpenFolder(string folderPath)
        {
            try
            {
                // 使用 Process.Start 启动文件资源管理器
                Process.Start("explorer.exe", folderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法打开文件夹: {ex.Message}");
            }
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
