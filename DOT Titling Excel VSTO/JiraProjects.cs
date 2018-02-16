using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class JiraProjects
    {
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Projects"))
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateTable(app, activeWorksheet);
                    }
                    else
                    {
                        MessageBox.Show("Missing Columns: " + missingColumns);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        //Update Table Data
        private static void UpdateTable(Excel.Application app, Excel.Worksheet ws)
        {
            try
            {
                var projects = GetAllFromJira().Result;
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
                    UpdateValues(ws, jiraFields, currentRow, project, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateValues(Excel.Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Jira.Project project, bool notFound)
        {
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                string type = jiraField.Type;
                string item = jiraField.Value;
                string formula = jiraField.Formula;
                int column = SSUtils.GetColumnFromHeader(activeWorksheet, columnHeader);
                if (type == "Standard")
                    SSUtils.SetCellValue(activeWorksheet, row, column, ExtractStandardValue(project, item), columnHeader);
                if (type == "Function")
                    SSUtils.SetCellValue(activeWorksheet, row, column, ExtractValueBasedOnFunction(project, item), columnHeader);
            }
        }

        //Get From Jira
        private async static Task<Jira.Project> GetSingleFromJira(string issueID)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var project = await ThisAddIn.GlobalJira.Projects.GetProjectAsync(issueID);
                return project;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private async static Task<List<Jira.Project>> GetAllFromJira()
        {
            try
            {
                var projects = await ThisAddIn.GlobalJira.Projects.GetProjectsAsync();
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
