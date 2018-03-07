using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Text;
using Jira = Atlassian.Jira;
using Excel = Microsoft.Office.Interop.Excel;

//// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
namespace DOT_Titling_Excel_VSTO
{
    class JiraShared
    {
        // Enums
        public enum ImportType
        {
            AllIssues = 1,
            StoriesAndBugsOnly = 2,
            EpicsOnly = 3,
            TasksOnly = 4,
            ChecklistTasksOnly = 5
        };
        
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Epics")
                {
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        ImportType importType = ImportType.StoriesAndBugsOnly;
                        if (ws.Name == "Epics")
                        {
                            idColumnName = "Epic ID";
                            importType = ImportType.EpicsOnly;
                        }
                        List<Jira.Issue> issues = UpdateTable(app, ws, listofProjects, importType, idColumnName);
                        if (issues != null)
                        {
                            AddNewRowsToTable(app, ws, issues, listofProjects, importType, idColumnName);
                            string dt = DateTime.Now.ToString("MM/dd/yyyy");
                            string val = ws.Name + " (Updated on " + dt + ")";
                            SSUtils.SetCellValue(ws, 1, 1, val, "Updated On");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Missing Columns: " + missingColumns);
                    }
                }
                else
                {
                    MessageBox.Show(ws.Name + " can't be updated.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteAddNewRowsToTable(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Epics")
                {
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        ImportType importType = ImportType.StoriesAndBugsOnly;
                        if (ws.Name == "Epics")
                        {
                            idColumnName = "Epic ID";
                            importType = ImportType.EpicsOnly;
                        }
                        AddNewRowsToTable(app, ws, null, listofProjects, importType, idColumnName);
                    }
                    else
                    {
                        MessageBox.Show("Missing Columns: " + missingColumns);
                    }
                }
                else
                {
                    MessageBox.Show(ws.Name + " can't be updated.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteUpdateSelectedRows(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Epics")
                {
                    var selection = app.Selection;
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        if (ws.Name == "Epics")
                            idColumnName = "Epic ID";
                        UpdateSelectedRows(app, ws, selection, idColumnName);
                    }
                    else
                    {
                        MessageBox.Show("Missing Columns: " + missingColumns);
                    }
                }
                else
                {
                    MessageBox.Show(ws.Name + " can't be updated.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static bool ExecuteSaveSelectedCellsToJira(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Epics")
                {
                    var selection = app.Selection;
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        if (ws.Name == "Epics")
                            idColumnName = "Epic ID";
                        return SaveSelectedCellsToJira(ws, selection, idColumnName);
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
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        public static void ExecuteUpdateRowBeforeOperation(Excel.Application app, Excel.Worksheet ws, string issueID, string idColumnName)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    UpdateRowBeforeOperation(app, ws, issueID, idColumnName);
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

        //Get From Jira
        public async static Task<List<Jira.Issue>> GetAllFromJira(List<string> listofProjects, ImportType importType)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project in (");
                jql.Append(FormatProjectList(listofProjects));
                jql.Append(")");
                string jqlIssueTypes = GetJQLForImportType(importType);
                jql.Append(jqlIssueTypes);
                jql.Append(" AND ");
                jql.Append("summary ~ \"!DELETE\"");
                List<Jira.Issue> filteredIssues = await Filter(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private static string GetJQLForImportType(ImportType importType)
        {
            string jql = string.Empty;
            switch (importType)
            {
                case ImportType.AllIssues:
                    break;
                case ImportType.StoriesAndBugsOnly:
                    jql = " AND issuetype in (\"Software Bug\", Story)";
                    break;
                case ImportType.EpicsOnly:
                    jql = " AND issuetype in (\"Epic\")";
                    break;
                case ImportType.TasksOnly:
                    jql = " AND issuetype in (\"Task\")";
                    break;
                case ImportType.ChecklistTasksOnly:
                    //TO DO
                    break;
                default:
                    break;
            }
            return jql;
        }

        public async static Task<Jira.Issue> GetSingleFromJira(string issueID)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var issue = await ThisAddIn.GlobalJira.Issues.GetIssueAsync(issueID);
                return issue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<List<Jira.Issue>> GetSelectedFromJira(List<string> listofIssueIDs)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("key in (");
                jql.Append(FormatListofIDs(listofIssueIDs));
                jql.Append(")");
                List<Jira.Issue> filteredIssues = await Filter(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private async static Task<IDictionary<string, Jira.Issue>> GetSelectedFromJiraAlternative(params string[] listofIssueIDs)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesAsync(listofIssueIDs);
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static async Task<List<Jira.Issue>> Filter(StringBuilder jql)
        {
            var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize);
            var totalIssues = issues.TotalItems;
            var totalPages = (double)totalIssues / (double)ThisAddIn.PageSize;
            totalPages = Math.Ceiling(totalPages);
            var allIssues = issues.ToList();
            for (int currentPage = 1; currentPage < totalPages; currentPage++)
            {
                int startRecord = ThisAddIn.PageSize * currentPage;
                issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize, startRecord);
                allIssues.AddRange(issues.ToList());
                if (issues.Count() == 0)
                {
                    break;
                }
            }
            var filteredIssues = allIssues.Where(i =>
                        i.Summary != "DELETE").ToList();
            return filteredIssues;
        }

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

        //Update Table Data
        public static List<Jira.Issue> UpdateTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType, string idColumnName)
        {
            try
            {
                var issues = GetAllFromJira(listofProjects, importType).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                int headerRow = SSUtils.GetHeaderRow(ws);
                int footerRow = SSUtils.GetFooterRow(ws);

                int projectKeyCol = SSUtils.GetColumnFromHeader(ws, "Project Key");
                int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string projectKey = SSUtils.GetCellValue(ws, currentRow, projectKeyCol);
                    if (listofProjects.Contains(projectKey))
                    {
                        string id = SSUtils.GetCellValue(ws, currentRow, idColumn);
                        var issue = issues.FirstOrDefault(i => i.Key == id);
                        UpdateRow(ws, jiraFields, currentRow, issue, issue != null);
                    }
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static void UpdateSelectedRows(Excel.Application app, Excel.Worksheet ws, Excel.Range selection, string idColumnName)
        {
            List<Jira.Issue> issues = GetListofSelectedIssuesIDsFromTable(ws, selection, idColumnName);
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
            int headerRow = SSUtils.GetHeaderRow(ws);
            int footerRow = SSUtils.GetFooterRow(ws);
            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                    string id = SSUtils.GetCellValue(ws, row, idColumn).Trim();
                    var issue = issues.FirstOrDefault(i => i.Key == id);
                    UpdateRow(ws, jiraFields, row, issue, issue != null);
                    SSUtils.SetStandardRowHeight(ws, row, row);
                }
            }
        }

        public static void AddNewRowsToTable(Excel.Application app, Excel.Worksheet ws, List<Jira.Issue> issues, List<string> listofProjects, ImportType importType, string idColumnName)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    if (issues == null)
                        issues = GetAllFromJira(listofProjects, importType).Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

                    List<string> listofIDs = new List<string>();
                    Excel.Range idColumnRange = ws.get_Range(wsRangeName + "[" + idColumnName + "]", Type.Missing);
                    foreach (Excel.Range cell in idColumnRange.Cells)
                    {
                        listofIDs.Add(cell.Value);
                    }
                    foreach (var id in listofIDs)
                    {
                        issues.Remove(issues.FirstOrDefault(x => x.Key.Value == id.ToString()));
                    }

                    string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                    foreach (var issue in issues)
                    {
                        Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                        int footerRow = footerRangeRange.Row;
                        Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                        rToInsert.Insert();
                        string status = GetStatus(ws, footerRow);
                        UpdateRow(ws, jiraFields, footerRow, issue, issue != null);
                        UpdateRowAfterAdd(app, ws, issue, footerRow, status);
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Rows Added.");
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

        public static void UpdateRow(Excel.Worksheet ws, List<JiraFields> jiraFields, int row, Jira.Issue issue, bool found)
        {
            try
            {
                foreach (var jiraField in jiraFields)
                {
                    string columnHeader = jiraField.ColumnHeader;
                    string type = jiraField.Type;
                    string item = jiraField.Value;
                    string formula = jiraField.Formula;
                    int column = SSUtils.GetColumnFromHeader(ws, columnHeader);
                    if (found)
                    {
                        if (type == "Standard")
                            SSUtils.SetCellValue(ws, row, column, ExtractStandardValue(issue, item), columnHeader);
                        if (type == "Custom")
                            SSUtils.SetCellValue(ws, row, column, ExtractCustomValue(issue, item), columnHeader);
                        if (type == "Function")
                            SSUtils.SetCellValue(ws, row, column, ExtractValueBasedOnFunction(issue, item), columnHeader);
                    }
                    else
                    {
                        if (item == "issue.Summary")
                        {
                            int issueTypeCol = SSUtils.GetColumnFromHeader(ws, "Issue Type");
                            if (issueTypeCol != 0)
                                SSUtils.SetCellValue(ws, row, issueTypeCol, "{DELETED}", columnHeader);
                        }
                        SSUtils.SetCellValue(ws, row, column, string.Empty, columnHeader);
                    }
                    if (type == "Formula")
                        SSUtils.SetCellFormula(ws, row, column, formula);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void UpdateRowBeforeOperation(Excel.Application app, Excel.Worksheet ws, string issueID, string idColumnName)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                var issue = GetSingleFromJira(issueID).Result;
                int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                int row = SSUtils.FindTextInColumn(ws, rangeName + "[" + idColumnName + "]", issueID);
                UpdateRow(ws, jiraFields, row, issue, issue != null);
                SSUtils.SetStandardRowHeight(ws, row, row);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateRowAfterAdd(Excel.Application app, Excel.Worksheet ws, Jira.Issue issue, int row, string previousStatus)
        {
            if (issue.Type.Name == "Story" || issue.Type.Name == "Software Bug")
            {
                //Issue ID
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Issue ID"), issue.Key.Value, "Issue ID");

                //Summary (Local)
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Summary (Local)"), issue.Summary, "Summary (Local)");
                
                //Release (Local)
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Release (Local)"), SSUtils.GetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Release")), "Release (Local)");
                
                //Epic (Local)
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                int epicColumn = SSUtils.GetColumnFromHeader(ws, "Epic");
                string newEpic = SSUtils.GetCellValue(ws, row, epicColumn);
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Epic (Local)"), newEpic, "Epic (Local)");
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                
                //Sprint Number (Local)
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Sprint Number (Local)"), SSUtils.GetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Sprint Number")), "Sprint Number (Local)");

                //Status (Last Changed)
                if (issue.Project == ThisAddIn.ProjectKeyDOT)
                {
                    string newStatus = GetStatus(ws, row);
                    int statusLastChangedColumn = SSUtils.GetColumnFromHeader(ws, "Status (Last Changed)");
                    if (statusLastChangedColumn != 0)
                    {
                        string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");
                        int sprintColumn = SSUtils.GetColumnFromHeader(ws, "DOT Sprint Number (Local)");
                        string sprint = SSUtils.GetCellValue(ws, row, sprintColumn);

                        if (sprint != currentSprint)
                        {
                            SSUtils.SetCellValue(ws, row, statusLastChangedColumn, string.Empty, "Status (Last Changed)");
                        }
                        else
                        {
                            if (newStatus == "Done" || newStatus == "Ready for Development" || newStatus == "")
                            {
                                SSUtils.SetCellValue(ws, row, statusLastChangedColumn, string.Empty, "Status (Last Changed)");
                            }
                            else
                            if (newStatus != previousStatus)
                            {
                                SSUtils.SetCellValue(ws, row, statusLastChangedColumn, DateTime.Now.ToString("MM/dd/yyyy"), "Status (Last Changed)");
                            }
                        }
                    }
                }
            }

            if (issue.Type.Name == "Epic")
            {
                //Epic ID
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Epic ID"), issue.Key.Value, "Epic ID");
                //Epic
                SSUtils.SetCellValue(ws, row, SSUtils.GetColumnFromHeader(ws, "Epic"), issue.Summary, "Epic");
            }
        }

        //Extract
        public static string ExtractRelease(Jira.Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.FixVersions)
            {
                if (c > 0)
                    val = val + "; ";
                val = val + issue.FixVersions[c].Name;
                c++;
            }
            return val;
        }

        public static string ExtractLabels(Jira.Issue issue)
        {
            string val = string.Empty;
            if (issue.Labels.Count > 0)
            {
                foreach (var label in issue.Labels)
                {
                    val = val + "[" + label + "]";
                }
            }
            return val;
        }

        public static List<string> ExtractListOfLabels(Jira.Issue issue)
        {
            List<string> listofLabels = new List<string>();
            if (issue.Labels.Count > 0)
            {
                foreach (var label in issue.Labels)
                {
                    listofLabels.Add(label);
                }
            }
            return listofLabels;
        }

        public static bool SaveCustomField(string issueID, string field, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetSingleFromJira(issueID).Result;
                if (issue[field] == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
                if (newValue == string.Empty)
                {
                    issue[field] = null;
                }
                else
                {
                    issue[field] = newValue;
                }
                issue.SaveChanges(); if (!multiple)
                    MessageBox.Show(field + " successfully updated.");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                //MessageBox.Show(field + " could NOT successfully updated.");
                return false;
            }
        }

        //Save to Jira
        public static bool SaveSelectedCellsToJira(Excel.Worksheet ws, Excel.Range selection, string idColumnName)
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

                        int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                        string id = SSUtils.GetCellValue(ws, row, idColumn);

                        string type = string.Empty;
                        if (idColumnName == "Issue ID")
                        { 
                            int typeCol = SSUtils.GetColumnFromHeader(ws, "Issue Type");
                            type = SSUtils.GetCellValue(ws, row, typeCol);
                        }
                        if (idColumnName == "Epic ID")
                        {
                            type = "Epic";
                        }

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
                                    SaveSummary(id, newValue, multiple);
                                    break;
                                case "Status":
                                    SaveStatus(id, newValue, multiple);
                                    break;
                                case "Bypass Approval":
                                    if (type == "Story")
                                    {
                                        newValue = newValue.Trim();
                                        if (newValue != string.Empty && newValue != "x")
                                        {
                                            MessageBox.Show(fieldToSave + " is not valid. Required blank of x. (" + row + ")");
                                            break;
                                        }
                                        SaveYesNo(id, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
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
                                        SaveCustomField(id, fieldToSave, newValue, multiple);
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
                                        SaveCustomField(id, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - As A":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(id, "Story: As a(n)", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - Id Like":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(id, "Story: I'd like to be able to", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - So That":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(id, "Story: So that I can", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story Points":
                                    SaveCustomField(id, "Story Points", newValue, multiple);
                                    break;
                                case "DOT Jira ID":
                                    if (type == "Software Bug")
                                    {
                                        SaveCustomField(id, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a Software Bug. (" + row + ")");
                                    }
                                    break;
                                case "Release":
                                    SaveRelease(id, newValue, multiple);
                                    break;
                                case "Labels":
                                    SaveLabels(id, newValue, multiple);
                                    break;
                                case "Epic Link":
                                    SaveCustomField(id, "Epic Link", newValue, multiple);
                                    break;
                                case "SWAG":
                                    SaveCustomField(id, "SWAG", newValue, multiple);
                                    break;
                                case "Reason Blocked or Delayed":
                                    SaveCustomField(id, "Reason Blocked or Delayed", newValue, multiple);
                                    break;
                                case "Sprint":
                                    //SaveCustomField(id, "Sprint", newValue, multiple);
                                    SaveSprint(id, newValue, multiple);
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

        //Save Single Value
        public static bool SaveSummary(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                if (issue.Summary == newValue)
                {
                    MessageBox.Show("No change needed.");
                    return true;
                }
                issue.Summary = newValue;
                issue.SaveChanges();
                if (!multiple)
                    MessageBox.Show("Summary updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Summary could NOT be successfully updated.");
                return false;
            }
        }

        //Save Single Value
        public static bool SaveYesNo(string issueID, string item, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                string yesNo = ExtractValueBasedOnFunction(issue, item);
                if (yesNo == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
                if (newValue == "x")
                    SaveCustomField(issueID, item, "Yes", multiple);
                if (newValue == "")
                    SaveCustomField(issueID, item, string.Empty, multiple);
                if (!multiple)
                    MessageBox.Show(item + " updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show(item + " could NOT be successfully updated.");
                return false;
            }
        }

        public static bool SaveRelease(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                string curRelease = ExtractRelease(issue);
                if (curRelease == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }

                // Remove all of the existing versions
                /// PWH VERSION
                var oldVersions = issue.FixVersions.ToList();
                foreach (var oldVersion in oldVersions)
                {
                    issue.FixVersions.Remove(oldVersion);
                }

                if (newValue.Trim() != string.Empty)
                    issue.FixVersions.Add(newValue);

                issue.SaveChanges();
                if (!multiple)
                    MessageBox.Show("Release updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        //PWH
        public static bool SaveSprint(string issueID, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetSingleFromJira(issueID).Result;
                if (issue["Sprint"] == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }

                //string projectID = issue.Project;
                //var project = JiraProjects.GetSingleFromJira(projectID);


                //foreach (var sprint in project.Result.Sp)
                //{
                //    val = val + " " + ver;
                //}


                //    var allIssues = _jira.Issues.Queryable
                //.Where(i => i.Project == AppSettings.TargetProjectName
                //            && i["Sprint"] == new LiteralMatch(sprintName))
                //.ToList();

                string id = issue["Sprint"].Value;


                if (newValue == string.Empty)
                {
                    issue["Sprint"] = null;
                }
                else
                {
                    Jira.LiteralMatch s = new Jira.LiteralMatch(newValue);
                    issue["Sprint"].Value = s.ToString();
                }
                issue.SaveChanges(); if (!multiple)
                    MessageBox.Show("Sprint successfully updated.");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                //MessageBox.Show("Sprint could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveLabels(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                List<string> listofJiraLabels = ExtractListOfLabels(issue);
                List<string> listofExcelLabels = CreateListOfLabels(newValue);
                List<string> addLabels = listofExcelLabels.Except(listofJiraLabels).ToList();
                List<string> removeLabels = listofJiraLabels.Except(listofExcelLabels).ToList();

                if (addLabels.Count > 0)
                {
                    foreach (string label in addLabels)
                    {
                        issue.Labels.Add(label);
                    }
                    issue.SaveChanges();
                }

                if (removeLabels.Count > 0)
                {
                    foreach (string label in removeLabels)
                    {
                        issue.Labels.Remove(label);
                    }
                    issue.SaveChanges();
                }
                return true;
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveStatus(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                if (issue.Status.Name == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
                issue.WorkflowTransitionAsync(newValue);
                if (!multiple)
                    MessageBox.Show("Status transitioned successfully.");
                return true;
            }
            catch
            {
                MessageBox.Show("Status could NOT be transitioned to " + newValue);
                return true;
            }
        }

        //Extraction Methods
        private static string ExtractStandardValue(Jira.Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "issue.Project":
                    val = issue.Project;
                    break;
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
                case "issue.Assignee":
                    val = issue.Assignee;
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

        private static string ExtractValueBasedOnFunction(Jira.Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "Sprint Number":
                    val = ExtractSprintNumber(issue);
                    break;
                case "Release":
                    val = ExtractRelease(issue);
                    break;
                case "DOT Web Services":
                    val = ExtractDOTWebServices(issue);
                    break;
                case "Bypass Approval":
                    val = ExtractYesNo(issue, item);
                    break;
                case "Labels":
                    List<string> listofLabels = ExtractListOfLabels(issue);
                    foreach (string label in listofLabels)
                    {
                        val = val + label + ", ";
                    }
                    if (val != string.Empty && val.Right(2) == ", ")
                        val = val.Left(val.Length - 2);
                    break;
                default:
                    break;
            }
            return val;
        }

        //Extraction Functions
        private static string ExtractDOTWebServices(Jira.Issue issue)
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

        private static string ExtractSprintNumber(Jira.Issue issue)
        {
            string val = string.Empty;
            int thisSprint = 0;
            int lastSprint = 0;
            foreach (var value in issue.CustomFields["Sprint"].Values)
            {
                val = value;
                if (val.Length > 2)
                {
                    val = val.Substring(val.Length - 3).Trim();
                    if (val != string.Empty)
                    {
                        if (Int32.TryParse(val, out thisSprint))
                        {
                            if (thisSprint > lastSprint)
                                lastSprint = thisSprint;
                        }
                    }
                }
            }

            string sprintNumber = string.Empty;
            if (lastSprint == 0)
            {
                sprintNumber = "";
            }
            else
            {
                sprintNumber = lastSprint.ToString();
            }
            return sprintNumber;
        }

        private static string ExtractYesNo(Jira.Issue issue, string item)
        {
            string retval = string.Empty;
            string yesNo = ExtractCustomValue(issue, item);
            if (yesNo == "Yes")
                retval = "x";
            return retval;
        }

        //Lists
        private static List<string> CreateListOfLabels(string labels)
        {
            labels = labels.Replace(", ", ",");
            return labels.Split(',').ToList();
        }

        public static List<Jira.Issue> GetListofSelectedIssuesIDsFromTable(Excel.Worksheet ws, Excel.Range selection, string idColumnName)
        {
            List<string> listofIssues = new List<string>();
            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int issueIDCol = SSUtils.GetColumnFromHeader(ws, idColumnName);
                    string issueID = SSUtils.GetCellValue(ws, row, issueIDCol).Trim();
                    listofIssues.Add(issueID);
                }
            }
            var issues = GetSelectedFromJira(listofIssues).Result;
            return issues;
        }

        //String Builders
        private static StringBuilder FormatProjectList(List<string> listofProjects)
        {
            var projectList = new StringBuilder();
            int cnt = 1;
            int projectCount = listofProjects.Count();
            foreach (string project in listofProjects)
            {
                projectList.Append(project);
                if (cnt != projectCount)
                    projectList.Append(", ");

                cnt++;
            }
            return projectList;
        }

        private static StringBuilder FormatListofIDs(List<string> lst)
        {
            var idList = new StringBuilder();
            int cnt = 1;
            int iCnt = lst.Count();
            foreach (string project in lst)
            {
                idList.Append(project);
                if (cnt != iCnt)
                    idList.Append(", ");
                cnt++;
            }
            return idList;
        }

        private static string GetStatus(Excel.Worksheet ws, int footerRow)
        {
            string status = string.Empty;
            int statusColumn = SSUtils.GetColumnFromHeader(ws, "Status");
            if (statusColumn != 0)
                status = SSUtils.GetCellValue(ws, footerRow, statusColumn);
            return status;
        }
    }
}
