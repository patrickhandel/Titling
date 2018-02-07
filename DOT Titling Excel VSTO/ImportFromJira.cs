using System;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class ImportFromJira
    {      
        public static void ExecuteUpdateIssues(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Issues") || (activeWorksheet.Name == "Program Issues") || (activeWorksheet.Name == "DOT Releases"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateAllIssues(app, activeWorksheet, listofProjects);
                        AddNewIssues(app, activeWorksheet, listofProjects);
                        string dt = DateTime.Now.ToString("MM/dd/yyyy");
                        string val = activeWorksheet.Name +" (Updated on " + dt + ")";
                        SSUtils.SetCellValue(activeWorksheet, 1, 1, val, "Updated On");
                        TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteUpdateSelectedIssues(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;
                string table = SSUtils.GetSelectedTable(app);
                if (activeCell != null && ((table == "IssueData") || (table == "DOTReleaseData")))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateSelectedIssues(app, activeWorksheet, selection, listofProjects);
                        //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteAddIssues(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Issues") || (activeWorksheet.Name == "DOT Releases"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        AddNewIssues(app, activeWorksheet, listofProjects);
                        //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteUpateIssueBeforeMailMerge(string issueID)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ws = app.Sheets["Issues"];
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    UpdateIssueBeforeMailMerge(app, ws, issueID);
                    //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewIssues(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = JiraIssue.GetAllStoriesAndBugs(listofProjects).Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int column = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                    List<string> listOfissueIDs = new List<string>();
                    Excel.Range issueIDColumnRange = ws.get_Range(wsRangeName + "[Issue ID]", Type.Missing);
                    foreach (Excel.Range cell in issueIDColumnRange.Cells)
                    {
                        listOfissueIDs.Add(cell.Value);
                    }
                    foreach (var issueID in listOfissueIDs)
                    {
                        issues.Remove(issues.FirstOrDefault(x => x.Key.Value == issueID.ToString()));
                    }

                    string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                    foreach (var issue in issues)
                    {
                        Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                        int footerRow = footerRangeRange.Row;
                        Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                        rToInsert.Insert();
                        UpdateIssueValues(ws, jiraFields, footerRow, issue, false);

                        //Issue ID (2)
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "Issue ID");

                        if (ws.Name == "Issues")
                        {
                            //Summary (5)
                            SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Summary (Local)"), issue.Summary, "Summary (Local)");

                            //TO DO FIX
                            SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release (Local)"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release")), "Release (Local)");

                            //Epic
                            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                            int epicColumn = SSUtils.GetColumnFromHeader(ws, "Epic");
                            string newEpic = SSUtils.GetCellValue(ws, footerRow, epicColumn);
                            SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic (Local)"), newEpic, "Epic (Local)");
                            app.Calculation = Excel.XlCalculation.xlCalculationManual;

                            //Sprint Number (Local)
                            SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Sprint Number (Local)"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Sprint Number")), "Sprint Number (Local)");
                        }
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Issues Added.");
                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteUpdateEpics_DOT(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Epics"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateEpics(app, activeWorksheet, listofProjects);
                        AddNewEpics(app, activeWorksheet, listofProjects);
                        TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Light);
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

        private static void AddNewEpics(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = JiraIssue.GetAllEpics(listofProjects).Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int column = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                    List<string> listOfissueIDs = new List<string>();
                    Excel.Range issueIDColumnRange = ws.get_Range(wsRangeName + "[Epic ID]", Type.Missing);
                    foreach (Excel.Range cell in issueIDColumnRange.Cells)
                    {
                        listOfissueIDs.Add(cell.Value);
                    }
                    foreach (var issueID in listOfissueIDs)
                    {
                        issues.Remove(issues.FirstOrDefault(x => x.Key.Value == issueID.ToString()));
                    }

                    string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                    foreach (var issue in issues)
                    {
                        Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                        int footerRow = footerRangeRange.Row;
                        Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                        rToInsert.Insert();
                        UpdateIssueValues(ws, jiraFields, footerRow, issue, false);
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "issue.Key.Value");
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic (Local)"), issue.Summary, "Summary (Local)");
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Epics Added.");

                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteUpdateChecklist(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Project Checklist"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateChecklistTasks(app, activeWorksheet, listofProjects);
                        AddNewChecklistTasks(app, activeWorksheet, listofProjects);
                        TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteUpdateProjects(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Projects"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateProjects(app, activeWorksheet);
                        TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Light);
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

        private static void AddNewChecklistTasks(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = JiraIssue.GetAllTasks(listofProjects).Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int column = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                    List<string> listOfissueIDs = new List<string>();
                    Excel.Range issueIDColumnRange = ws.get_Range(wsRangeName + "[Issue ID]", Type.Missing);
                    foreach (Excel.Range cell in issueIDColumnRange.Cells)
                    {
                        listOfissueIDs.Add(cell.Value);
                    }
                    foreach (var issueID in listOfissueIDs)
                    {
                        issues.Remove(issues.FirstOrDefault(x => x.Key.Value == issueID.ToString()));
                    }

                    string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                    foreach (var issue in issues)
                    {
                        Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                        int footerRow = footerRangeRange.Row;
                        Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0", footerRow), Type.Missing);
                        rToInsert.Insert();
                        UpdateIssueValues(ws, jiraFields, footerRow, issue, false);
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "issue.Key.Value");
                        //SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic (Local)"), issue.Summary, "Summary (Local)");
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Tasks Added.");

                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateIssueBeforeMailMerge(Excel.Application app, Excel.Worksheet ws, string issueID)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                var issue  = JiraIssue.GetIssue(issueID).Result;
                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                int row = SSUtils.FindTextInColumn(ws, rangeName + "[Issue ID]", issueID);
                bool notFound = issue == null;
                UpdateIssueValues(ws, jiraFields, row, issue, notFound);
                SSUtils.SetStandardRowHeight(ws, row, row);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateEpics(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                var epics = JiraIssue.GetAllEpics(listofProjects).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                int cnt = epics.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string issueID = SSUtils.GetCellValue(ws, currentRow, issueIDCol);
                    var issue = epics.FirstOrDefault(i => i.Key == issueID);
                    bool notFound = issue == null;
                    UpdateIssueValues(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateChecklistTasks(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                var tasks = JiraIssue.GetAllTasks(listofProjects).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                int cnt = tasks.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string jiraID = SSUtils.GetCellValue(ws, currentRow, issueIDCol);
                    var issue = tasks.FirstOrDefault(i => i.Key == jiraID);
                    bool notFound = issue == null;
                    UpdateIssueValues(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateProjects(Excel.Application app, Excel.Worksheet ws)
        {
            try
            {
                var projects = JiraProject.GetAllProjects().Result;
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
                    UpdateProjectValues(ws, jiraFields, currentRow, project, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateAllIssues(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
            try
            {
                var issues = JiraIssue.GetAllStoriesAndBugs(listofProjects).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                int cnt = issues.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string jiraID = SSUtils.GetCellValue(ws, currentRow, issueIDCol);
                    var issue = issues.FirstOrDefault(i => i.Key == jiraID);
                    bool notFound = issue == null;
                    UpdateIssueValues(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateSelectedIssues(Excel.Application app, Excel.Worksheet ws, Excel.Range selection, List<string> listofProjects)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

            var issues = JiraIssue.GetAllStoriesAndBugs(listofProjects).Result;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

            string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
            Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            int cnt = selection.Rows.Count;

            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                    string issueID = SSUtils.GetCellValue(ws, row, issueIDCol).Trim();
                    if (issueID.Length > 10 && issueID.Substring(0, 10) == listofProjects[0] + "-")
                    {
                        var issue = issues.FirstOrDefault(p => p.Key.Value == issueID);
                        bool notFound = issue == null;
                        UpdateIssueValues(ws, jiraFields, row, issue, notFound);
                        SSUtils.SetStandardRowHeight(ws, row, row);
                    }
                }
            }
        }

        private static void UpdateIssueValues(Excel.Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Jira.Issue issue, bool notFound)
        {
            try
            {
                //Get the current status
                int statusColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "Status");            
                string newStatus = string.Empty;
                string previousStatus = string.Empty;
                if (statusColumn != 0)
                    previousStatus = SSUtils.GetCellValue(activeWorksheet, row, statusColumn);

                foreach (var jiraField in jiraFields)
                {
                    string columnHeader = jiraField.ColumnHeader;
                    string type = jiraField.Type;
                    string item = jiraField.Value;
                    string formula = jiraField.Formula;
                    int column = SSUtils.GetColumnFromHeader(activeWorksheet, columnHeader);

                    if (notFound)
                    {
                        string valueToSave = string.Empty;
                        if (item == "issue.Summary")
                        {
                            valueToSave = "{DELETED}";
                            int issueTypeCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Issue Type");
                            if (issueTypeCol != 0)
                                SSUtils.SetCellValue(activeWorksheet, row, issueTypeCol, valueToSave, columnHeader);
                        }
                        SSUtils.SetCellValue(activeWorksheet, row, column, valueToSave, columnHeader);
                    }
                    else
                    {
                        if (type == "Standard")
                            SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractStandardValue(issue, item), columnHeader);
                        if (type == "Custom")
                            SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractCustomValue(issue, item), columnHeader);
                        if (type == "Function")
                            SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractValueBasedOnFunction(issue, item), columnHeader);
                    }
                    if (type == "Formula")
                        SSUtils.SetCellFormula(activeWorksheet, row, column, formula);
                }

                if (activeWorksheet.Name == "Issues" && statusColumn != 0)
                {
                    newStatus = SSUtils.GetCellValue(activeWorksheet, row, statusColumn);
                    int statusLastChangedColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "Status (Last Changed)");
                    if (statusLastChangedColumn != 0)
                    {
                        string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");
                        int sprintColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "DOT Sprint Number (Local)");
                        string sprint = SSUtils.GetCellValue(activeWorksheet, row, sprintColumn);

                        if (sprint != currentSprint)
                        {
                            SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, string.Empty, "Status (Last Changed)");
                        }
                        else
                        {
                            if (newStatus == "Done" || newStatus == "Ready for Development" || newStatus == "")
                            {
                                SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, string.Empty, "Status (Last Changed)");
                            }
                            else
                            if (newStatus != previousStatus)
                            {
                                SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, DateTime.Now.ToString("MM/dd/yyyy"), "Status (Last Changed)");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateProjectValues(Excel.Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Jira.Project project, bool notFound)
        {
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                string type = jiraField.Type;
                string item = jiraField.Value;
                string formula = jiraField.Formula;
                int column = SSUtils.GetColumnFromHeader(activeWorksheet, columnHeader);
                if (type == "Standard")
                    SSUtils.SetCellValue(activeWorksheet, row, column, JiraProject.ExtractStandardValue(project, item), columnHeader);
                if (type == "Function")
                    SSUtils.SetCellValue(activeWorksheet, row, column, JiraProject.ExtractValueBasedOnFunction(project, item), columnHeader);
            }
        }

        public static string MissingColumns(Excel.Worksheet ws)
        {
            string missingFields = string.Empty;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                if (SSUtils.GetColumnFromHeader(ws, columnHeader) == 0)
                    missingFields = missingFields + ' ' + columnHeader;
            }
            return missingFields.Trim();
        }

        private static void UpdateSprintProgress(Excel.Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Jira.Issue issue, bool notFound)
        {
            // Get the Current Date
            string dt = DateTime.Now.ToString("MM/dd/yyyy");

            // Get the Current Sprint
            string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");

            //Get List of Issues
            List<Issue> issues = Lists.GetListOfIssues(activeWorksheet);
            issues = issues.FindAll(r => r.Sprint == currentSprint && r.Type == "Story");
            
            //TO DO
            //Add each issue to the bottom of the Table
        }
    }
}
