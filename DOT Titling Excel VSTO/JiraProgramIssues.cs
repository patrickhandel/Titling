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
    class JiraProgramIssues
    {
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if (activeWorksheet.Name == "Program Issues")
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateTable(app, activeWorksheet, listofProjects);
                        AddNewRowsToTable(app, activeWorksheet, listofProjects);
                        string dt = DateTime.Now.ToString("MM/dd/yyyy");
                        string val = activeWorksheet.Name + " (Updated on " + dt + ")";
                        SSUtils.SetCellValue(activeWorksheet, 1, 1, val, "Updated On");
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

        public static void ExecuteAddNewRowsToTable(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if (activeWorksheet.Name == "Program Issues")
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        AddNewRowsToTable(app, activeWorksheet, listofProjects);
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

        public static void ExecuteUpdateSelectedRows(Excel.Application app, List<string> listofProjects)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;
                string table = SSUtils.GetSelectedTable(app);
                if (activeCell != null && table == "ProgramIssueData")
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateSelectedRows(app, activeWorksheet, selection, listofProjects);
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

        public static bool ExecuteSaveSelectedCellsToJira(Excel.Application app)
        {
            try
            {
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                var selection = app.Selection;
                if (activeCell != null && activeWorksheet.Name == "Program Issues" )
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
        private static void UpdateTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
            try
            {
                var issues = GetAllFromJira(listofProjects).Result;
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
                    UpdateRow(ws, jiraFields, currentRow, issue, notFound);
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
            ////Summary (5)
            //SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Summary (Local)"), issue.Summary, "Summary (Local)");

            ////TO DO FIX
            //SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release (Local)"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release")), "Release (Local)");

            ////Epic
            //app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            //int epicColumn = SSUtils.GetColumnFromHeader(ws, "Epic");
            //string newEpic = SSUtils.GetCellValue(ws, footerRow, epicColumn);
            //SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic (Local)"), newEpic, "Epic (Local)");
            //app.Calculation = Excel.XlCalculation.xlCalculationManual;

            ////Sprint Number (Local)
            //SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Sprint Number (Local)"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Sprint Number")), "Sprint Number (Local)");
        }

        private static void UpdateSelectedRows(Excel.Application app, Excel.Worksheet ws, Excel.Range selection, List<string> listofProjects)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

            var issues = GetAllFromJira(listofProjects).Result;
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
                        UpdateRow(ws, jiraFields, row, issue, notFound);
                        SSUtils.SetStandardRowHeight(ws, row, row);
                    }
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
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
                jql.Append("summary ~ \"!DELETE\"");
                List<Jira.Issue> filteredIssues = await JiraShared.Filter(jql);
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
                                case "Story Points":
                                    SaveCustomField(issueID, "Story Points", newValue, multiple);
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

        //Extraction Functions

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