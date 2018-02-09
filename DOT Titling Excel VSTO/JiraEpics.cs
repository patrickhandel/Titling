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
    class JiraEpics
    {
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Epics"))
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateTable(app, activeWorksheet, listofProjects);
                        AddNewRowsToTable(app, activeWorksheet, listofProjects);
                        TableStandardization.Execute(app, TableStandardization.StandardizationType.Light);
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
        private static void UpdateTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
            try
            {
                var epics = GetAllFromJira(listofProjects).Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                int headerRow = SSUtils.GetHeaderRow(ws);
                int footerRow = SSUtils.GetFooterRow(ws);
                int projectKeyCol = SSUtils.GetColumnFromHeader(ws, "Project Key");
                int epicIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string projectKey = SSUtils.GetCellValue(ws, currentRow, projectKeyCol);
                    if (listofProjects.Contains(projectKey))
                    {
                        string issueID = SSUtils.GetCellValue(ws, currentRow, epicIDCol);
                        var epic = epics.FirstOrDefault(i => i.Key == issueID);
                        bool notFound = false;
                        if (epic == null)
                            notFound = true;
                        UpdateRow(ws, jiraFields, currentRow, epic, notFound);
                    }
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewRowsToTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = GetAllFromJira(listofProjects).Result;
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
                        UpdateRow(ws, jiraFields, footerRow, issue, false);
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "issue.Key.Value");
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), issue.Summary, "Epic");
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
                if (activeCell != null && activeWorksheet.Name == "Epics")
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
        
        //Get From Jira
        private async static Task<List<Jira.Issue>> GetAllFromJira(List<string> listofProjects)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project in (");
                jql.Append(JiraShared.FormatProjectList(listofProjects));
                jql.Append(")");
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Epic\")");
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

        private static async Task<List<Jira.Issue>> Filter(StringBuilder jql)
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

        //Save to Jira
        private static bool SaveSelectedCellsToJira(Excel.Worksheet ws, Excel.Range activeCell)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);

                int column = activeCell.Column;
                int row = activeCell.Row;
                string fieldToSave = SSUtils.GetCellValue(ws, headerRowRange.Row, column);
                string newValue = SSUtils.GetCellValue(ws, row, column).Trim();

                int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                string issueID = SSUtils.GetCellValue(ws, row, issueIDCol);
                bool multiple = false;
                switch (fieldToSave)
                {
                    case "Summary":
                        SaveSummary(issueID, newValue, multiple);
                        break;
                    case "Status":
                        SaveStatus(issueID, newValue, multiple);
                        break;
                    case "Story Points":
                        SaveCustomField(issueID, "Story Points", newValue, multiple);
                        break;
                    default:
                        MessageBox.Show(fieldToSave + " can't be updated.");
                        break;
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
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
                case "Release":
                    val = JiraShared.ExtractRelease(issue);
                    break;
                case "Labels":
                    List<string> listofLabels = JiraShared.ExtractListOfLabels(issue);
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

        //Save Single Value
        private static bool SaveSummary(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = JiraShared.GetSingleFromJira(issueID).Result;
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

        private static bool SaveRelease(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = JiraShared.GetSingleFromJira(issueID).Result;
                string curRelease = JiraShared.ExtractRelease(issue);
                if (curRelease == newValue)
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
                    MessageBox.Show("Release updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        private static bool SaveLabels(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = JiraShared.GetSingleFromJira(issueID).Result;
                List<string> listofJiraLabels = JiraShared.ExtractListOfLabels(issue);
                List<string> listofExcelLabels = JiraShared.CreateListOfLabels(newValue);
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

        private static bool SaveStatus(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = JiraShared.GetSingleFromJira(issueID).Result;
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

        private static bool SaveCustomField(string issueID, string field, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = JiraShared.GetSingleFromJira(issueID).Result;
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
    }
}
