﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class ImportData
    {
        //Public Methods
        public async static Task<bool> ExecuteUpdateTable(Jira.Jira jira, Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if ((ws.Name == "Projects"))
                {
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        bool success = await UpdateTable(jira, app, ws);
                        return success;
                    }
                    else
                    {
                        MessageBox.Show("Missing Columns: " + missingColumns);
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show(ws.Name + " can't be updated.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        //Update Table Data
        private async static Task<bool> UpdateTable(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws)
        {
            try
            {
                var projects = GetAllFromJira(jira).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                int cnt = projects.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int keyCol = SSUtils.GetColumnFromHeader(ws, "Project Key");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string key = SSUtils.GetCellValue(ws, currentRow, keyCol);
                    var project = projects.FirstOrDefault(i => i.Key == key);
                    bool notFound = projects == null;
                    bool success = await UpdateValues(ws, jiraFields, currentRow, project, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private async static Task<bool> UpdateValues(Excel.Worksheet ws, List<JiraFields> jiraFields, int row, Jira.Project project, bool notFound)
        {
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                string type = jiraField.Type;
                string item = jiraField.Value;
                string formula = jiraField.Formula;
                int column = SSUtils.GetColumnFromHeader(ws, columnHeader);
                if (type == "Standard")
                    SSUtils.SetCellValue(ws, row, column, ExtractStandardValue(project, item));
                if (type == "Function")
                    SSUtils.SetCellValue(ws, row, column, ExtractValueBasedOnFunction(project, item));
            }
            return true;
        }

        //Get From Jira
        public async static Task<Jira.Project> GetSingleFromJira(Jira.Jira jira, string issueID)
        {
            try
            {
                jira.Issues.MaxIssuesPerRequest = 1;
                var project = await jira.Projects.GetProjectAsync(issueID);
                return project;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private async static Task<List<Jira.Project>> GetAllFromJira(Jira.Jira jira)
        {
            try
            {
                var projects = await jira.Projects.GetProjectsAsync();
                return projects.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        //Save to Jira

        //Save Single Value

        //Extraction Methods
        private static string ExtractStandardValue(Jira.Project project, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "Project.Name":
                    val = project.Name;
                    break;
                case "Project.Key":
                    val = project.Key;
                    break;
                case "Project.Id":
                    val = project.Id;
                    break;
                case "Project.Lead":
                    val = project.Lead;
                    break;
                default:
                    break;
            }
            return val;
        }

        private static string ExtractCustomValue(Jira.Issue issue, string item)
        {
            string val = string.Empty;
            item = item.Replace(" Id ", " I'd ");
            item = item.Trim();
            try
            {
                val = issue[item].Value;
            }
            catch
            {
                val = string.Empty;
            }
            return val;
        }

        private static string ExtractValueBasedOnFunction(Jira.Project project, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                default:
                    break;
            }
            return val;
        }

        //Extraction Functions

        //Save Single Value
    }
}
