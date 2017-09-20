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
    class Import
    {
        public static void ExecuteAddNewTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.DoStandardStuff(app);
                Worksheet ws = app.Sheets["Tickets"];
                AddNewTickets(app, ws);
                WorksheetStandardization.ExecuteCleanup(ws);
                SSUtils.DoStandardStuff(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewTickets(Excel.Application app, Worksheet ws)
        {
            try
            {
                List<Issue> issues = GetAllTicketsFromJira().ToList();
                string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                int column = SSUtils.GetColumnFromHeader(ws, "Story ID");
                List<JiraFields> jiraFields = WorksheetPropertiesManager.GetJiraFields("TicketData");

                List<string> listOfStoryIDs = new List<string>();
                Range storyIDColumnRange = ws.get_Range(wsRangeName + "[Story ID]", Type.Missing);
                foreach (Range cell in storyIDColumnRange.Cells)
                {
                    listOfStoryIDs.Add(cell.Value);
                }
                foreach (var storyID in listOfStoryIDs)
                {
                    issues.Remove(issues.FirstOrDefault(x => x.Key.Value == storyID.ToString()));
                }

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                foreach (var issue in issues)
                {
                    Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                    int footerRow = footerRangeRange.Row;
                    Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                    rToInsert.Insert();
                    UpdateValues(ws, jiraFields, footerRow, issue, false);
                    SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value);
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Summary"), issue.Summary);
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Story Release"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Story Release")));
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Epic")));
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Hufflepuff Sprint"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Hufflepuff Sprint")));
                    SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                }
                MessageBox.Show(issues.Count() + " Tickets Added.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteImportAllJiraTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;

                if (activeCell != null && (activeWorksheet.Name == "Tickets"))
                {
                    SSUtils.DoStandardStuff(app);
                    ImportAllJiraTickets(app, activeWorksheet, selection);
                    WorksheetStandardization.ExecuteCleanup(activeWorksheet);
                    SSUtils.DoStandardStuff(app);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteImportSelectedJiraTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;

                if (activeCell != null && (activeWorksheet.Name == "Tickets"))
                {
                    app.ScreenUpdating = false;
                    app.Calculation = XlCalculation.xlCalculationManual;
                    ImportSelectedJiraTickets(app, activeWorksheet, selection);
                    WorksheetStandardization.ExecuteCleanup(activeWorksheet);
                    app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteImportSingleJiraTicket(string jiraId)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Worksheet ws = app.Sheets["Tickets"];
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;
                ImportSingleJiraTicket(app, ws, jiraId);
                WorksheetStandardization.ExecuteCleanup(ws);
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void ImportSingleJiraTicket(Excel.Application app, Excel.Worksheet ws, string jiraId)
        {
            try
            {
                List<JiraFields> jiraFields = WorksheetPropertiesManager.GetJiraFields("TicketData");
                List<Issue> issues = GetSingleTicketFromJira(jiraId);

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);

                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Story ID");
                int row = SSUtils.FindTextInColumn(ws, "TicketData[Story ID]", jiraId);

                foreach (var issue in issues)
                {
                    bool notFound = issue == null;
                    UpdateValues(ws, jiraFields, row, issue, notFound);
                    SSUtils.SetStandardRowHeight(ws, row, row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void ImportAllJiraTickets(Excel.Application app, Excel.Worksheet activeWorksheet, Excel.Range selection)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

            List<Issue> issues = GetAllTicketsFromJira();
            List<JiraFields> jiraFields = WorksheetPropertiesManager.GetJiraFields("TicketData");

            int cnt = issues.Count();

            string sHeaderRangeName = SSUtils.GetHeaderRangeName(activeWorksheet.Name);
            Range headerRowRange = activeWorksheet.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            string sFooterRowRange = SSUtils.GetFooterRangeName(activeWorksheet.Name);
            Range footerRangeRange = activeWorksheet.get_Range(sFooterRowRange, Type.Missing);
            int footerRow = footerRangeRange.Row;

            Range rToInsert = activeWorksheet.get_Range(String.Format("{0}:{1}", footerRow, footerRow + cnt - 1), Type.Missing);
            Range rToDelete = activeWorksheet.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);

            rToInsert.Insert();
            rToDelete.Delete();

            int jiraIDCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Story ID");
            int row = headerRow + 1;
            foreach (var issue in issues)
            {
                bool notFound = issue == null;
                UpdateValues(activeWorksheet, jiraFields, row, issue, notFound);
                SSUtils.SetStandardRowHeight(activeWorksheet, row, row);
                row++;
            }
        }

        private static void ImportSelectedJiraTickets(Excel.Application app, Worksheet ws, Range selection)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

            List<Issue> issues = GetAllTicketsFromJira();

            string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
            List<JiraFields> jiraFields = WorksheetPropertiesManager.GetJiraFields(rangeName);

            string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
            Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            int cnt = selection.Rows.Count;

            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Story ID");
                    string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol).Trim();
                    if (jiraId.Length > 10 && jiraId.Substring(0, 10) == "DOTTITLNG-")
                    {
                        var issue = issues.FirstOrDefault(p => p.Key.Value == jiraId);
                        bool notFound = issue == null;
                        UpdateValues(ws, jiraFields, row, issue, notFound);
                        SSUtils.SetStandardRowHeight(ws, row, row);
                    }
                }
            }
        }

        private static void UpdateValues(Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Issue issue, bool notFound)
        {
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                string type = jiraField.Type;
                string value = jiraField.Value;
                string formula = jiraField.Formula;
                int column = SSUtils.GetColumnFromHeader(activeWorksheet, columnHeader);
                
                if (notFound)
                {
                    string valueToSave = string.Empty;
                    if (value == "issue.Summary")
                        valueToSave = "{DELETED}";
                    SSUtils.SetCellValue(activeWorksheet, row, column, valueToSave);
                }
                else
                {
                    if (type == "Standard")
                        SSUtils.SetCellValue(activeWorksheet, row, column, GetStandardIssueValueForCell(issue, value));
                    if (type == "Custom")
                        SSUtils.SetCellValue(activeWorksheet, row, column, GetCustomIssueValueForCell(issue, value));
                    if (type == "Function")
                        SSUtils.SetCellValue(activeWorksheet, row, column, GetFunctionIssueValueForCell(issue, value));
                }
                if (type == "Formula")
                    SSUtils.SetCellFormula(activeWorksheet, row, column, formula);
            }
        }

        private static List<Issue> GetAllTicketsFromJira()
        {
            var jira = Jira.CreateRestClient(ThisAddIn.JiraSite, ThisAddIn.JiraUserName, ThisAddIn.JiraPassword);
            jira.MaxIssuesPerRequest = 1000;
            var issues = (from i in jira.Issues.Queryable
                          where i.Project == "DOTTITLNG" &&
                          (i.Type == "Story" || i.Type == "Software Bug") &&
                          i.Summary != "DELETE"
                          orderby i.Created
                          select i).ToList();

            // try this: i.Summary == new LiteralMatch("My Title")

            var issuesToRemove = issues.FindAll(x => x.Summary.ToUpper().Trim() == "DELETE");
            foreach (var issueToRemove in issuesToRemove)
            {
                issues.Remove(issues.FirstOrDefault(x => x.Key.Value == issueToRemove.Key.Value));
            }
            return issues;
        }

        private static List<Issue> GetSingleTicketFromJira(string jiraId)
        {
            try
            {
                var jira = Jira.CreateRestClient(ThisAddIn.JiraSite, ThisAddIn.JiraUserName, ThisAddIn.JiraPassword);
                jira.MaxIssuesPerRequest = 1000;
                var issues = (from i in jira.Issues.Queryable
                              where i.Key == jiraId
                              select i).ToList();
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private static string GetStandardIssueValueForCell(Issue issue, string value)
        { 
            string val = string.Empty;
            switch (value)
            {
                case "issue.Type.Name":
                    val = issue.Type.Name;
                    break;
                case "issue.Key.Value":
                    val = issue.Key.Value;
                    break;
                case "issue.Summary":
                    val = issue.Summary;
                    break;
                case "issue.Status.Name":
                    val = issue.Status.Name;
                    break;
                case "issue.Description":
                    val = issue.Description;
                    break;
                default:
                    break;
            }
            return val;
        }

        private static string GetFunctionIssueValueForCell(Issue issue, string value)
        {
            string val = string.Empty;
            switch (value)
            {
                case "Sprint":
                    val = ExtractSprintNumber(issue);
                    break;
                case "Release":
                    val = ExtractRelease(issue);
                    break;
                case "Fix Release":
                    val = ExtractFixRelease(issue);
                    break;
                case "DOT Web Services":
                    val = ExtractDOTWebServices(issue);
                    break;
                default:
                    break;
            }
            return val;
        }


        private static string GetCustomIssueValueForCell(Issue issue, string value)
        {
            string val = string.Empty;
            value = value.Replace(" Id ", " I'd ");
            try
            {
                val = issue[value].Value;
            }
            catch
            {
                val = string.Empty;
            }
            return val;
        }

        private static string ExtractRelease(Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.AffectsVersions)
            {
                val = issue.AffectsVersions[c].Name;
                c++;
            }
            return val;
        }

        private static string ExtractFixRelease(Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.FixVersions)
            {
                val = issue.FixVersions[c].Name;
                c++;
            }
            return val;
        }

        private static string ExtractDOTWebServices(Issue issue)
        {
            string val = string.Empty;
            if (issue["DOT Web Services"] != null)
            { 
                foreach (var ver in issue.CustomFields["DOT Web Services"].Values)
                {
                    val = val + " " + ver;
                }
                val = val.Trim().Replace(" ", ", ");
            }
            return val;
        }

        private static string ExtractSprintNumber(Issue issue)
        {
            string val = GetCustomIssueValueForCell(issue, "Sprint");
            if (val != string.Empty)
            {
                val = string.Empty;
                foreach (var value in issue.CustomFields["Sprint"].Values)
                    val = value;
                val = val.Replace("DOT", "");
                val = val.Replace("Backlog", "");
                val = val.Replace("Hufflepuff", "");
                val = val.Replace("Sprint", "");
                val = val.Replace("Ready", "");
                val = val.Replace("Other", "");
                val = val.Replace("Approved", "");
                val = val.Replace("-", "");
                val = val.Replace(" ", "");
                for (int rev = 1; rev <= 12; rev++)
                    val = val.Replace("R" + rev.ToString(), "");
            }
            return val;
        }
    }
}
