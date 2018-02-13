using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class JiraIssues : JiraShared
    {
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app, List<string> listofProjects, ImportType importType)
        {
            try
            {
                var wsIssues = app.Sheets["Issues"];
                wsIssues.Select();
                string missingColumns = SSUtils.MissingColumns(wsIssues);
                if (missingColumns == string.Empty)
                {
                    UpdateTable(app, wsIssues, listofProjects, importType);
                    AddNewRowsToTable(app, wsIssues, listofProjects, importType);
                    string dt = DateTime.Now.ToString("MM/dd/yyyy");
                    string val = wsIssues.Name + " (Updated on " + dt + ")";
                    SSUtils.SetCellValue(wsIssues, 1, 1, val, "Updated On");
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

        public static void ExecuteAddNewRowsToTable(Excel.Application app, List<string> listofProjects, ImportType importType)
        {
            try
            {
                var wsIssues = app.Sheets["Issues"];
                wsIssues.Select();
                string missingColumns = SSUtils.MissingColumns(wsIssues);
                if (missingColumns == string.Empty)
                {
                    AddNewRowsToTable(app, wsIssues, listofProjects, importType);
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

        public static void ExecuteUpdateSelectedRows(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;
                string table = SSUtils.GetSelectedTable(app);
                if (activeCell != null && ((table == "IssueData") || (table == "DOTReleaseData")))
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateSelectedRows(app, activeWorksheet, selection);
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

        public static void ExecuteUpdateRowBeforeOperation(string issueID)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ws = app.Sheets["Issues"];
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    UpdateRowBeforeOperation(app, ws, issueID);
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

        public static bool ExecuteSaveSelectedCellsToJira(Excel.Application app)
        {
            try
            {
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                var selection = app.Selection;
                if (activeCell != null && (activeWorksheet.Name == "Issues" || activeWorksheet.Name == "DOT Releases"))
                {
                    return SaveSelectedCellsToJira(activeWorksheet, activeCell);
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        //Update Table Data
        private static void UpdateTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType)
        {
            try
            {
                var issues = GetAllFromJira(listofProjects, importType).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                int headerRow = SSUtils.GetHeaderRow(ws);
                int footerRow = SSUtils.GetFooterRow(ws);
                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string issueID = SSUtils.GetCellValue(ws, currentRow, issueIDCol);
                    var issue = issues.FirstOrDefault(i => i.Key == issueID);
                    bool notFound = false;
                    if (issue == null)
                        notFound = true;
                    UpdateRow(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewRowsToTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = GetAllFromJira(listofProjects, importType).Result;
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
                        UpdateRow(ws, jiraFields, footerRow, issue, false);

                        //Issue ID (2)
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "Issue ID");

                        UpdateRowAfterAdd(app, ws, issue, footerRow);
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

        private static void UpdateRowAfterAdd(Excel.Application app, Excel.Worksheet ws, Jira.Issue issue, int footerRow)
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

        private static void UpdateSelectedRows(Excel.Application app, Excel.Worksheet ws, Excel.Range selection)
        {
            List<Jira.Issue> issues = GetListofSelectedIssuesIDsFromTable(ws, selection);
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
            int headerRow = SSUtils.GetHeaderRow(ws);
            int footerRow = SSUtils.GetFooterRow(ws);
            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                    string issueID = SSUtils.GetCellValue(ws, row, issueIDCol).Trim();
                    var issue = issues.FirstOrDefault(i => i.Key == issueID);
                    bool notFound = false;
                    if (issue == null)
                        notFound = true;
                    UpdateRow(ws, jiraFields, row, issue, notFound);
                    int column = SSUtils.GetColumnFromHeader(ws, "Project Key");
                    SSUtils.SetCellValue(ws, footerRow, column, issue.Project, "Project Key");
                    SSUtils.SetStandardRowHeight(ws, row, row);
                }
            }
        }

        private static void UpdateRow(Excel.Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Jira.Issue issue, bool notFound)
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
                            SSUtils.SetCellValue(activeWorksheet, row, column, ExtractStandardValue(issue, item), columnHeader);
                        if (type == "Custom")
                            SSUtils.SetCellValue(activeWorksheet, row, column, ExtractCustomValue(issue, item), columnHeader);
                        if (type == "Function")
                            SSUtils.SetCellValue(activeWorksheet, row, column, ExtractValueBasedOnFunction(issue, item), columnHeader);
                    }
                    if (type == "Formula")
                        SSUtils.SetCellFormula(activeWorksheet, row, column, formula);
                }

                if (notFound == false && issue.Project == ThisAddIn.ProjectKeyDOT)
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

        private static void UpdateRowBeforeOperation(Excel.Application app, Excel.Worksheet ws, string issueID)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                var issue = GetSingleFromJira(issueID).Result;
                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                int row = SSUtils.FindTextInColumn(ws, rangeName + "[Issue ID]", issueID);
                bool notFound = false;
                if (issue == null)
                    notFound = true;
                UpdateRow(ws, jiraFields, row, issue, notFound);
                SSUtils.SetStandardRowHeight(ws, row, row);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        //Get From Jira
        private async static Task<List<Jira.Issue>> GetAllTasksFromJira(List<string> listofProjects)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project = " + listofProjects[0]);
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Task\")");
                jql.Append(" AND ");
                jql.Append("\"Epic Link\" = " + listofProjects[0] + "-945");
                List<Jira.Issue> filteredIssues = await Filter(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        //Save to Jira
        private static bool SaveSelectedCellsToJira(Excel.Worksheet ws, Excel.Range selection)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);

                int cellCount = selection.Cells.Count;
                bool multiple = (cellCount > 1);
                foreach (Excel.Range cell in selection.Cells)
                {
                    int row = cell.Row;
                    if (ws.Rows[row].EntireRow.Height != 0)
                    {
                        int column = cell.Column;
                        string fieldToSave = SSUtils.GetCellValue(ws, headerRowRange.Row, column);
                        string newValue = SSUtils.GetCellValue(ws, row, column).Trim();

                        int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                        string issueID = SSUtils.GetCellValue(ws, row, issueIDCol);

                        int typeCol = SSUtils.GetColumnFromHeader(ws, "Issue Type");
                        string type = SSUtils.GetCellValue(ws, row, typeCol);

                        int summaryCol = SSUtils.GetColumnFromHeader(ws, "Summary");
                        string summary = SSUtils.GetCellValue(ws, row, summaryCol);

                        if (summary == "{DELETED}")
                        {
                            MessageBox.Show("Cannot update a Deleted issue.");
                        }
                        else
                        {
                            switch (fieldToSave)
                            {
                                case "Summary":
                                    SaveSummary(issueID, newValue, multiple);
                                    break;
                                case "Status":
                                    SaveStatus(issueID, newValue, multiple);
                                    break;
                                case "Date Submitted to DOT":
                                    if (type == "Story")
                                    {
                                        newValue = newValue.Trim();
                                        if (newValue != string.Empty)
                                        {
                                            if (SSUtils.CheckDate(newValue) == false)
                                            {
                                                MessageBox.Show(fieldToSave + " is not a valid date. (" + row + ")");
                                                break;
                                            }
                                            DateTime dt = DateTime.Parse(newValue);
                                            newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                        }
                                        SaveCustomField(issueID, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Date Approved by DOT":
                                    if (type == "Story")
                                    {
                                        newValue = newValue.Trim();
                                        if (newValue != string.Empty)
                                        {
                                            if (SSUtils.CheckDate(newValue) == false)
                                            {
                                                MessageBox.Show(fieldToSave + " is not a valid date. (" + row + ")");
                                                break;
                                            }
                                            DateTime dt = DateTime.Parse(newValue);
                                            newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                        }
                                        SaveCustomField(issueID, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - As A":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(issueID, "Story: As a(n)", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - Id Like":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(issueID, "Story: I'd like to be able to", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - So That":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(issueID, "Story: So that I can", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story Points":
                                    SaveCustomField(issueID, "Story Points", newValue, multiple);
                                    break;
                                case "DOT Jira ID":
                                    if (type == "Software Bug")
                                    {
                                        SaveCustomField(issueID, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a Software Bug. (" + row + ")");
                                    }
                                    break;
                                case "Release":
                                    SaveRelease(issueID, newValue, multiple);
                                    break;
                                case "Labels":
                                    SaveLabels(issueID, newValue, multiple);
                                    break;
                                case "Epic Link":
                                    SaveCustomField(issueID, "Epic Link", newValue, multiple);
                                    break;
                                case "SWAG":
                                    SaveCustomField(issueID, "SWAG", newValue, multiple);
                                    break;
                                case "Reason Blocked or Delayed":
                                    SaveCustomField(issueID, "Reason Blocked or Delayed", newValue, multiple);
                                    break;
                                default:
                                    MessageBox.Show(fieldToSave + " can't be updated. (" + row + ")");
                                    break;
                            }
                        }
                    }
                }
                return multiple;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return true;
            }
        }
    }
}
