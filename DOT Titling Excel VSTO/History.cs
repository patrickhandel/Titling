using System;
using System.Windows.Forms;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Atlassian.Jira;
using System.Collections.Generic;

namespace DOT_Titling_Excel_VSTO
{
    class History
    {
        public static void ExecuteGetHistory()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                if (activeWorksheet.Name == "Hist")
                {
                    GetHistory(app, activeWorksheet);
                    WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public async static void GetHistory(Excel.Application app, Worksheet ws)
        {
            ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
            int x = 1;
            var issues = JiraUtils.GetAllIssues().Result;
            foreach (var i in issues)
            {
                var historyItems = await ThisAddIn.GlobalJira.Issues.GetChangeLogsAsync(i.Key.Value);
                foreach (var item in historyItems)
                {
                    string author = item.Author.ToString();
                    SSUtils.SetCellValue(ws, x, 1, "DOTTITLNG-165");
                    SSUtils.SetCellValue(ws, x, 2, author);
                    x++;
                }
            }
        }
    }
}
