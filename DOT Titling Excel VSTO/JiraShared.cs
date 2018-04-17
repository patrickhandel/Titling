using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

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

        public async static Task<Jira.Jira> GetJira(Excel.Application app)
        {
            try
            {
                string jiraUserName = app.get_Range("JiraUserName").Value2;
                string jiraPassword = app.get_Range("JiraPassword").Value2;
                Jira.Jira jira = Jira.Jira.CreateRestClient(ThisAddIn.JiraSite, jiraUserName, jiraPassword);
                return jira;
            }
            //catch (Exception ex)
            catch
            {
                MessageBox.Show("Error : Not a properly formatted workbook.");
                return null;
            }
        }

        //Public Methods
        public async static Task<bool> ExecuteUpdateTable(Jira.Jira jira, Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;

                string tableRangeName = SSUtils.GetSelectedTable(app);
                Int32 rowCount = SSUtils.TableRowCount(ws, tableRangeName);
                if (rowCount <= ThisAddIn.MaxRecordsToProcess)
                {
                    if (ws.Name == "Issues" || ws.Name == "Program Issues" || ws.Name == "Epics")
                    {
                        string missingColumns = SSUtils.MissingColumns(ws);
                        if (missingColumns == string.Empty)
                        {
                            string idColumnName = "Issue ID";
                            ImportType importType = ImportType.StoriesAndBugsOnly;
                            if (ws.Name == "Program Issues")
                            {
                                importType = ImportType.AllIssues;
                            }
                            if (ws.Name == "Epics")
                            {
                                idColumnName = "Epic ID";
                                importType = ImportType.EpicsOnly;
                            }
                            List<Jira.Issue> issues = await UpdateTable(jira, app, ws, listofProjects, importType, idColumnName);
                            if (issues != null)
                            {
                                bool success = await AddNewRowsToTable(jira, app, ws, issues, listofProjects, importType, idColumnName);
                                string dt = DateTime.Now.ToString("MM/dd/yyyy");
                                string val = ws.Name + " (Updated on " + dt + ")";
                                SSUtils.SetCellValue(ws, 1, 1, val);
                                return success;
                            }
                            else
                            {
                                return false;
                            }
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
                else
                {
                    MessageBox.Show("Cannot update more than " + ThisAddIn.MaxRecordsToProcess + " rows. Please select few rows.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        public async static Task<bool> ExecuteAddNewRowsToTable(Jira.Jira jira, Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Program Issues" || ws.Name == "Epics")
                {
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        ImportType importType = ImportType.StoriesAndBugsOnly;
                        if (ws.Name == "Progam Issues")
                        {
                            importType = ImportType.AllIssues;
                        }
                        if (ws.Name == "Epics")
                        {
                            idColumnName = "Epic ID";
                            importType = ImportType.EpicsOnly;
                        }
                        bool success = await AddNewRowsToTable(jira, app, ws, null, listofProjects, importType, idColumnName);
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

        public async static Task<bool> ExecuteUpdateSelectedRows(Jira.Jira jira, Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Program Issues" || ws.Name == "Epics")
                {
                    Excel.Range selection = app.Selection;
                    Int32 rowCount = SSUtils.TableSelectedRowCount(ws, selection);
                    if (rowCount <= ThisAddIn.MaxRecordsToProcess)
                    {
                        string missingColumns = SSUtils.MissingColumns(ws);
                        if (missingColumns == string.Empty)
                        {
                            string idColumnName = "Issue ID";
                            if (ws.Name == "Epics")
                                idColumnName = "Epic ID";
                            bool success = await UpdateSelectedRows(jira, app, ws, selection, idColumnName);
                        }
                        else
                        {
                            MessageBox.Show("Missing Columns: " + missingColumns);
                            return false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Cannot update more than " + ThisAddIn.MaxRecordsToProcess + " rows. Please select few rows.");
                        return false;
                    }
                    return true;
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

        public async static Task<bool> ExecuteSaveSelectedCellsToJira(Jira.Jira jira, Excel.Application app, List<string> listofProjects)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if (ws.Name == "Issues" || ws.Name == "Program Issues" || ws.Name == "Epics")
                {
                    Excel.Range selection = app.Selection;
                    string missingColumns = SSUtils.MissingColumns(ws);
                    if (missingColumns == string.Empty)
                    {
                        string idColumnName = "Issue ID";
                        if (ws.Name == "Epics")
                            idColumnName = "Epic ID";
                        bool multiple = await SaveSelectedCellsToJira(jira, ws, selection, idColumnName);
                        return multiple;
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

        public async static Task<bool> ExecuteUpdateRowBeforeOperation(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, string issueID, string idColumnName)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    bool success = await UpdateRowBeforeOperation(jira, app, ws, issueID, idColumnName);
                    return success;
                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        //Get From Jira
        public async static Task<List<Jira.Issue>> GetAllFromJira(Jira.Jira jira, List<string> listofProjects, ImportType importType)
        {
            try
            {
                jira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project in (");
                jql.Append(FormatProjectList(listofProjects));
                jql.Append(")");
                string jqlIssueTypes = GetJQLForImportType(importType);
                jql.Append(jqlIssueTypes);
                jql.Append(" AND ");
                jql.Append("summary ~ \"!DELETE\"");

                int totalItems = await GetIssueCountFromJira(jira, jql);

                List<Jira.Issue> filteredIssues = await Filter(jira, jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<List<Jira.Issue>> GetSelectedFromJira(Jira.Jira jira, List<string> listofIssueIDs)
        {
            try
            {
                jira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("key in (");
                jql.Append(FormatListofIDs(listofIssueIDs));
                jql.Append(")");

                int totalItems = await GetIssueCountFromJira(jira, jql);

                List<Jira.Issue> filteredIssues = await Filter(jira, jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private async static Task<List<Jira.Issue>> GetAllTasksFromJira(Jira.Jira jira, List<string> listofProjects)
        {
            try
            {
                jira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project = " + listofProjects[0]);
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Task\")");
                jql.Append(" AND ");
                jql.Append("\"Epic Link\" = " + listofProjects[0] + "-945");

                int totalItems = await GetIssueCountFromJira(jira, jql);

                List<Jira.Issue> filteredIssues = await Filter(jira, jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<Jira.Issue> GetSingleFromJira(Jira.Jira jira, string issueID)
        {
            try
            {
                jira.Issues.MaxIssuesPerRequest = 1;
                var issue = await jira.Issues.GetIssueAsync(issueID);
                return issue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static async Task<List<Jira.Issue>> Filter(Jira.Jira jira, StringBuilder jql)
        {
            var issues = await jira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize);
            var totalIssues = issues.TotalItems;
            var totalPages = (double)totalIssues / (double)ThisAddIn.PageSize;
            totalPages = Math.Ceiling(totalPages);
            var allIssues = issues.ToList();
            for (int currentPage = 1; currentPage < totalPages; currentPage++)
            {
                int startRecord = ThisAddIn.PageSize * currentPage;
                issues = await jira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize, startRecord);
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

        private static async Task<int> GetIssueCountFromJira(Jira.Jira jira, StringBuilder jql)
        {
            var issues = await jira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize);
            return issues.TotalItems;
        }

        private static string GetJQLForImportType(ImportType importType)
        {
            string jql = string.Empty;
            switch (importType)
            {
                case ImportType.AllIssues:
                    break;
                case ImportType.StoriesAndBugsOnly:
                    jql = " AND issuetype in (\"Software Bug\", \"Bug\", Story)";
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

        //Update Table Data
        public async static Task<List<Jira.Issue>> UpdateTable(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType, string idColumnName)
        {
            try
            {
                var issues = GetAllFromJira(jira, listofProjects, importType).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                int headerRow = SSUtils.GetHeaderRow(ws);
                int footerRow = SSUtils.GetFooterRow(ws);

                int projectKeyCol = SSUtils.GetColumnFromHeader(ws, "Project Key");
                int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string projectKey = SSUtils.GetCellValue(ws, row, projectKeyCol);
                    if (listofProjects.Contains(projectKey))
                    {
                        string id = SSUtils.GetCellValue(ws, row, idColumn);
                        var issue = issues.FirstOrDefault(i => i.Key == id);
                        bool success = await UpdateRow(ws, jiraFields, row, issue, issue != null);                        
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

        public static async Task<bool> UpdateSelectedRows(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, Excel.Range selection, string idColumnName)
        {
            List<Jira.Issue> issues = GetListofSelectedIssuesIDsFromTable(jira, ws, selection, idColumnName);
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
                    bool success = await UpdateRow(ws, jiraFields, row, issue, issue != null);
                    SSUtils.SetStandardRowHeight(ws, row, row);
                }
            }
            return true;
        }

        public static Int32 GetNumberOfIssuesToAdd(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, List<Jira.Issue> issues, List<string> listofProjects, ImportType importType, string idColumnName)
        {
            try
            {
                if (issues == null)
                    issues = GetAllFromJira(jira, listofProjects, importType).Result;
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
                return 999;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return 0;
            }
        }

        public static async Task<bool> AddNewRowsToTable(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, List<Jira.Issue> issues, List<string> listofProjects, ImportType importType, string idColumnName)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    if (issues == null)
                        issues = GetAllFromJira(jira, listofProjects, importType).Result;
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
                        bool success = await UpdateRow(ws, jiraFields, footerRow, issue, issue != null);                        
                        UpdateRowAfterAdd(app, ws, issue, footerRow, status);
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Rows Added.");
                    return true;
                }
                else
                {
                    MessageBox.Show("Missing Columns: " + missingColumns);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        public static async Task<bool> UpdateRow(Excel.Worksheet ws, List<JiraFields> jiraFields, int row, Jira.Issue issue, bool found)
        {
            try
            {
                bool isDeleted = false;
                int issueTypeCol = SSUtils.GetColumnFromHeader(ws, "Issue Type");
                if (issueTypeCol != 0)
                {
                    string issueType = SSUtils.GetCellValue(ws, row, issueTypeCol);
                    if (issueType == "{DELETED}")
                        isDeleted = true;
                }
                if (isDeleted == false)
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
                                SSUtils.SetCellValue(ws, row, column, ExtractStandardValue(issue, item));
                            if (type == "Custom")
                                SSUtils.SetCellValue(ws, row, column, ExtractCustomValue(issue, item));
                            if (type == "Function")
                                SSUtils.SetCellValue(ws, row, column, ExtractValueBasedOnFunction(issue, item));
                        }
                        else
                        {
                            if (item == "issue.Summary")
                            {
                                if (issueTypeCol != 0)
                                    SSUtils.SetCellValue(ws, row, issueTypeCol, "{DELETED}");
                            }
                            SSUtils.SetCellValue(ws, row, column, string.Empty);
                        }
                        if (type == "Formula")
                            SSUtils.SetCellFormula(ws, row, column, formula);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        public async static Task<bool> UpdateRowBeforeOperation(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, string issueID, string idColumnName)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                var issue = GetSingleFromJira(jira, issueID).Result;
                int idColumn = SSUtils.GetColumnFromHeader(ws, idColumnName);
                int row = SSUtils.FindTextInColumn(ws, rangeName + "[" + idColumnName + "]", issueID);
                bool success = await UpdateRow(ws, jiraFields, row, issue, issue != null);
                SSUtils.SetStandardRowHeight(ws, row, row);
                return success;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private static void UpdateRowAfterAdd(Excel.Application app, Excel.Worksheet ws, Jira.Issue issue, int row, string previousStatus)
        {
            if (ws.Name == "Issues" | ws.Name == "Program Issues")
            {
                //Issue ID
                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                if (issueIDCol != 0)
                    SSUtils.SetCellValue(ws, row, issueIDCol, issue.Key.Value);

                //Summary (Local)
                int summaryLocalColumn = SSUtils.GetColumnFromHeader(ws, "Summary (Local)");
                if (summaryLocalColumn != 0)
                    SSUtils.SetCellValue(ws, row, summaryLocalColumn, issue.Summary);

                //Release (Local)
                int releaseLocalColumn = SSUtils.GetColumnFromHeader(ws, "Release (Local)");
                int fixVersionColumn = SSUtils.GetColumnFromHeader(ws, "Fix Version");
                if (releaseLocalColumn != 0 && fixVersionColumn != 0)
                {
                    string fixVersion = SSUtils.GetCellValue(ws, row, fixVersionColumn);
                    SSUtils.SetCellValue(ws, row, releaseLocalColumn, fixVersion);
                }

                //Epic (Local)
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                int epicLocalColumn = SSUtils.GetColumnFromHeader(ws, "Epic (Local)");
                int epicColumn = SSUtils.GetColumnFromHeader(ws, "Epic");
                if (epicLocalColumn != 0 && epicColumn != 0)
                {
                    string newEpic = SSUtils.GetCellValue(ws, row, epicColumn);
                    SSUtils.SetCellValue(ws, row, epicLocalColumn, newEpic);
                }
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
            }

            if (ws.Name == "Issues")
            {
                //Sprint Number (Local)
                int sprintNumberLocalCol = SSUtils.GetColumnFromHeader(ws, "Sprint Number (Local)");
                int sprintNumberCol = SSUtils.GetColumnFromHeader(ws, "Sprint Number (Local)");
                if (sprintNumberLocalCol != 0 && sprintNumberCol != 0)
                {
                    string sprintNumber = SSUtils.GetCellValue(ws, row, sprintNumberCol);
                    SSUtils.SetCellValue(ws, row, sprintNumberLocalCol, sprintNumber);
                }

                //Status (Last Changed)
                if (issue.Project == ThisAddIn.ProjectKeyDOT)
                {
                    string newStatus = GetStatus(ws, row);
                    int statusLastChangedCol = SSUtils.GetColumnFromHeader(ws, "Status (Last Changed)");
                    if (statusLastChangedCol != 0)
                    {
                        string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");
                        int sprintColumn = SSUtils.GetColumnFromHeader(ws, "DOT Sprint Number (Local)");
                        string sprint = SSUtils.GetCellValue(ws, row, sprintColumn);
                        if (sprint != currentSprint)
                        {
                            SSUtils.SetCellValue(ws, row, statusLastChangedCol, string.Empty);
                        }
                        else
                        {
                            if (newStatus == "Done" || newStatus == "Ready for Development" || newStatus == "")
                            {
                                SSUtils.SetCellValue(ws, row, statusLastChangedCol, string.Empty);
                            }
                            else
                            if (newStatus != previousStatus)
                            {
                                SSUtils.SetCellValue(ws, row, statusLastChangedCol, DateTime.Now.ToString("MM/dd/yyyy"));
                            }
                        }
                    }
                }
            }

            if (ws.Name == "Epics")
            {
                //Epic
                int epicCol = SSUtils.GetColumnFromHeader(ws, "Epic");
                if (epicCol != 0)
                    SSUtils.SetCellValue(ws, row, epicCol, issue.Summary);

                //Epic ID
                int epicIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                if (epicIDCol != 0)
                    SSUtils.SetCellValue(ws, row, epicIDCol, issue.Key.Value);
            }
        }

        //Extract
        public static string ExtractFixVersion(Jira.Issue issue)
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

        public static string ExtractAffectsVersion(Jira.Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.AffectsVersions)
            {
                if (c > 0)
                    val = val + "; ";
                val = val + issue.AffectsVersions[c].Name;
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

        public static bool SaveCustomField(Jira.Jira jira, string issueID, string field, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetSingleFromJira(jira, issueID).Result;
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
                issue.SaveChanges();
                if (!multiple)
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
        public async static Task<bool> SaveSelectedCellsToJira(Jira.Jira jira, Excel.Worksheet ws, Excel.Range selection, string idColumnName)
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
                                    SaveSummary(jira, id, newValue, multiple);
                                    break;
                                case "Status":
                                    SaveStatus(jira, id, newValue, multiple);
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
                                        SaveYesNo(jira, id, fieldToSave, newValue, multiple);
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
                                        SaveCustomField(jira, id, fieldToSave, newValue, multiple);
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
                                        SaveCustomField(jira, id, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - As A":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(jira, id, "Story: As a(n)", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - Id Like":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(jira, id, "Story: I'd like to be able to", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - So That":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(jira, id, "Story: So that I can", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Release Notes":
                                    if (type == "Story")
                                    {
                                        SaveCustomField(jira, id, "Release Notes", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story Points":
                                    SaveCustomField(jira, id, "Story Points", newValue, multiple);
                                    break;
                                case "DOT Jira ID":
                                    if (type == "Software Bug" || type == "Bug")
                                    {
                                        SaveCustomField(jira, id, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a bBug. (" + row + ")");
                                    }
                                    break;
                                case "Fix Version":
                                    SaveFixVersion(jira, id, newValue, multiple);
                                    break;
                                case "Affects Version":
                                    SaveAffectsVersion(jira, id, newValue, multiple);
                                    break;
                                case "Labels":
                                    SaveLabels(jira, id, newValue, multiple);
                                    break;
                                case "Epic Link":
                                    SaveCustomField(jira, id, "Epic Link", newValue, multiple);
                                    break;
                                case "SWAG":
                                    SaveCustomField(jira, id, "SWAG", newValue, multiple);
                                    break;
                                case "Reason Blocked or Delayed":
                                    SaveCustomField(jira, id, "Reason Blocked or Delayed", newValue, multiple);
                                    break;
                                //case "Sprint":
                                //    //SaveCustomField(id, "Sprint", newValue, multiple);
                                //    SaveSprint(jira, id, newValue, multiple);
                                //    break;
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
        public static bool SaveSummary(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
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
        public static bool SaveYesNo(Jira.Jira jira, string issueID, string item, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
                string yesNo = ExtractValueBasedOnFunction(issue, item);
                if (yesNo == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
                if (newValue == "x")
                    SaveCustomField(jira, issueID, item, "Yes", multiple);
                if (newValue == "")
                    SaveCustomField(jira, issueID, item, string.Empty, multiple);
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

        public static bool SaveFixVersion(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
                string curVersion = ExtractFixVersion(issue);
                if (curVersion == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }

                // Remove all of the existing versions
                var oldVersions = issue.FixVersions.ToList();
                foreach (var oldVersion in oldVersions)
                {
                    issue.FixVersions.Remove(oldVersion);
                }

                if (newValue.Trim() != string.Empty)
                    issue.FixVersions.Add(newValue);

                issue.SaveChanges();
                if (!multiple)
                    MessageBox.Show("Fix Version updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Fix Version could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveAffectsVersion(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
                string curVersion = ExtractAffectsVersion(issue);
                if (curVersion == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }

                // Remove all of the existing versions
                var oldVersions = issue.AffectsVersions.ToList();
                foreach (var oldVersion in oldVersions)
                {
                    issue.AffectsVersions.Remove(oldVersion);
                }

                if (newValue.Trim() != string.Empty)
                    issue.AffectsVersions.Add(newValue);

                issue.SaveChanges();
                if (!multiple)
                    MessageBox.Show("Affects Version updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Affects Version could NOT successfully updated.");
                return false;
            }
        }


        public static bool SaveSprint(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetSingleFromJira(jira, issueID).Result;
                if (issue["Sprint"] == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
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
                issue.SaveChanges();
                if (!multiple)
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

        public static bool SaveLabels(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
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
                MessageBox.Show("Labels could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveStatus(Jira.Jira jira, string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(jira, issueID).Result;
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
                case "Fix Version":
                    val = ExtractFixVersion(issue);
                    break;
                case "Affects Version":
                    val = ExtractAffectsVersion(issue);
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

        public static List<Jira.Issue> GetListofSelectedIssuesIDsFromTable(Jira.Jira jira, Excel.Worksheet ws, Excel.Range selection, string idColumnName)
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
            var issues = GetSelectedFromJira(jira, listofIssues).Result;
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
