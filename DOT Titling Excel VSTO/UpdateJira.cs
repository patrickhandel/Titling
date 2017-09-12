using System;
using System.Windows.Forms;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Atlassian.Jira;
using System.Data;
using System.Collections.Generic;

namespace DOT_Titling_Excel_VSTO
{
    class UpdateJira
    {
        public static void ExecuteUpateSummary()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                if (activeCell != null && activeWorksheet.Name == "Tickets")
                {
                    SSUtils.DoStandardStuff(app);
                    UpdateSummaryAsync(activeWorksheet, activeCell);
                    SSUtils.DoStandardStuff(app);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static async void UpdateSummaryAsync(Excel.Worksheet ws, Excel.Range activeCell)
        {
            try
            {
                var jira = Jira.CreateRestClient(ThisAddIn.JiraSite, ThisAddIn.JiraUserName, ThisAddIn.JiraPassword);

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Story ID");
                int column = activeCell.Column;
                int row = activeCell.Row;
                string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);
                string newSummary = SSUtils.GetCellValue(ws, row, column);
                var issue = await jira.Issues.GetIssueAsync(jiraId);
                issue.Summary = newSummary;
                issue.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
