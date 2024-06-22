using System;
using System.Collections.Generic;
using Newtonsoft.Json;
//using RestSharp;
//using RestSharp.Authenticators;
using System.Windows.Forms;
using System.Net;
using System.Dynamic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools;
using Mspj = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;
using System.Net.NetworkInformation;
using System.IO;
using System.Linq;
//using static PjAddIn.JiraResponse_GetIssue_IssueLinks;
//using Microsoft.Office.Interop.MSProject;

namespace Pj2Excel
{
    static class Utils
    {
        //获取相对路径
        public static string GetPath()
        {
            // 使用 VSTO 获取当前打开的 MS Project 应用程序
            Microsoft.Office.Interop.MSProject.Application msProjectApp = Globals.ThisAddIn.Application;

            // 获取当前打开的项目
            Microsoft.Office.Interop.MSProject.Project activeProject = msProjectApp.ActiveProject;

            // 获取当前打开项目的文件路径
            string msProjectFilePath = activeProject.FullName;

            // 如果文件路径不为空，保存 Excel 文件到相同路径
            if (!string.IsNullOrEmpty(msProjectFilePath))
            {
                // 使用 Path 类来获取 MS Project 文件所在目录
                string projectDirectory = Path.GetDirectoryName(msProjectFilePath);

                // 构建相对路径
                //string relativePath = Path.Combine(projectDirectory, "MemberExcelTask_Interop.xlsx");

                // 保存 Excel 文件到相对路径
                //string excelFilePath = projectDirectory;
                //Console.WriteLine($"Excel 文件已成功生成：{excelFilePath}");

                // 返回相对路径
                return projectDirectory;
            }
            else
            {
                Console.WriteLine("无法确定 MS Project 文件路径。");
                return null;
            }
        }

        /// <summary>
        /// 获取所有选中的Ms Project中的Tasks
        /// </summary>
        /// <returns></returns>
        public static Mspj.Tasks GetActiveTasks()
        {
            Mspj.Project thisPj = Globals.ThisAddIn.Application.ActiveProject;
            Mspj.Tasks activeTasks = thisPj.Application.ActiveSelection.Tasks;
            return activeTasks;
        }
       
    }
}