using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class JiraEpics : JiraShared
    {
        //Public Methods
        public static void ExecuteUpdateTable(Excel.Application app, List<string> listofProjects, ImportType importType)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Epics"))
                {
                    string missingColumns = SSUtils.MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateTable(app, activeWorksheet, listofProjects, importType);
                        AddNewRowsToTable(app, activeWorksheet, listofProjects, importType);
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
        private static void UpdateTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType)
        {
            try
            {
                var epics = GetAllFromJira(listofProjects, importType).Result;
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

        private static void AddNewRowsToTable(Excel.Application app, Excel.Worksheet ws, List<string> listofProjects, ImportType importType)
        {
            try
            {
                string missingColumns = SSUtils.MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = GetAllFromJira(listofProjects, importType).Result;
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
    }
}








