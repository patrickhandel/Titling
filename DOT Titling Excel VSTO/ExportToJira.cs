using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class ExportToJira
    {
        public static bool ExecuteSaveSelectedTicketValues(Excel.Application app)
        {
            try
            {
                Worksheet activeWorksheet = app.ActiveSheet;
                Range activeCell = app.ActiveCell;
                var selection = app.Selection;
                if (activeCell != null && (activeWorksheet.Name == "Tickets" || activeWorksheet.Name == "DOT Releases"))
                {
                    return SaveSelectedTicketValues(activeWorksheet, selection);
                }
                if (activeCell != null && activeWorksheet.Name == "Epics")
                {
                    return SaveSelectedEpic(activeWorksheet, activeCell);
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private static bool SaveSelectedEpic(Worksheet ws, Range activeCell)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);

                int column = activeCell.Column;
                int row = activeCell.Row;
                string fieldToSave = SSUtils.GetCellValue(ws, headerRowRange.Row, column);
                string newValue = SSUtils.GetCellValue(ws, row, column).Trim();

                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);
                bool multiple = false;
                switch (fieldToSave)
                {
                    case "Jira Epic Summary":
                        JiraIssue.SaveSummary(jiraId, newValue, multiple);
                        break;
                    case "Jira Status":
                        JiraIssue.SaveStatus(jiraId, newValue, multiple);
                        break;
                    case "Jira Epic Points":
                        JiraIssue.SaveCustomField(jiraId, "Story Points", newValue, multiple);
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

        private static bool SaveSelectedTicketValues(Worksheet ws, Range selection)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);

                int cellCount = selection.Cells.Count;
                bool multiple = (cellCount > 1);
                foreach (Range cell in selection.Cells)
                {
                    int row = cell.Row;
                    if (ws.Rows[row].EntireRow.Height != 0)
                    {
                        int column = cell.Column;
                        string fieldToSave = SSUtils.GetCellValue(ws, headerRowRange.Row, column);
                        string newValue = SSUtils.GetCellValue(ws, row, column).Trim();

                        int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                        string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);

                        int typeCol = SSUtils.GetColumnFromHeader(ws, "Ticket Type");
                        string type = SSUtils.GetCellValue(ws, row, typeCol);

                        int summaryCol = SSUtils.GetColumnFromHeader(ws, "Jira Summary");
                        string summary = SSUtils.GetCellValue(ws, row, summaryCol);

                        if (summary == "{DELETED}")
                        {
                            MessageBox.Show("Cannot update a Deleted ticket.");
                        }
                        else
                        {
                            switch (fieldToSave)
                            {
                                case "Jira Summary":
                                    JiraIssue.SaveSummary(jiraId, newValue, multiple);
                                    break;
                                case "Jira Status":
                                    JiraIssue.SaveStatus(jiraId, newValue, multiple);
                                    break;
                                case "Date Submitted to DOT":
                                    if (type == "Story")
                                    {
                                        newValue = newValue.Trim();
                                        if (newValue != string.Empty)
                                        {
                                            if (CheckDate(newValue) == false)
                                            {
                                                MessageBox.Show(fieldToSave + " is not a valid date. (" + row + ")");
                                                break;
                                            }
                                            DateTime dt = DateTime.Parse(newValue);
                                            newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                        }
                                        JiraIssue.SaveCustomField(jiraId, fieldToSave, newValue, multiple);
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
                                            if (CheckDate(newValue) == false)
                                            {
                                                MessageBox.Show(fieldToSave + " is not a valid date. (" + row + ")");
                                                break;
                                            }
                                            DateTime dt = DateTime.Parse(newValue);
                                            newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                        }
                                        JiraIssue.SaveCustomField(jiraId, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - As A":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(jiraId, "Story: As a(n)", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - Id Like":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(jiraId, "Story: I'd like to be able to", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - So That":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(jiraId, "Story: So that I can", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Points":
                                    JiraIssue.SaveCustomField(jiraId, "Story Points", newValue, multiple);
                                    break;
                                case "DOT Jira ID":
                                    if (type == "Software Bug")
                                    {
                                        JiraIssue.SaveCustomField(jiraId, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a Software Bug. (" + row + ")");
                                    }
                                    break;
                                case "Jira Release":
                                    JiraIssue.SaveRelease(jiraId, newValue, multiple);
                                    break;
                                case "Labels":
                                    JiraIssue.SaveLabels(jiraId, newValue, multiple);
                                    break;
                                case "Jira Epic ID":
                                    JiraIssue.SaveCustomField(jiraId, "Epic Link", newValue, multiple);
                                    break;
                                case "SWAG":
                                    JiraIssue.SaveCustomField(jiraId, "SWAG", newValue, multiple);
                                    break;
                                case "Reason Blocked or Delayed":
                                    JiraIssue.SaveCustomField(jiraId, "Reason Blocked or Delayed", newValue, multiple);
                                    break;
                                //case "Backlog Area":
                                //    JiraUtils.SaveCustomField(jiraId, "Sprint", newValue);
                                //    break;
                                default:
                                    //DO NOT UPDATE THE FOLLOWING:
                                    //Ticket Type
                                    //Jira Fix Release
                                    //Jira Hufflepuff Sprint
                                    //Jira Epic
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

        private static bool CheckDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public static class DataTypeExtensions
    {
        #region Methods

        public static string Left(this string str, int length)
        {
            str = (str ?? string.Empty);
            return str.Substring(0, Math.Min(length, str.Length));
        }

        public static string Right(this string str, int length)
        {
            str = (str ?? string.Empty);
            return (str.Length >= length)
                ? str.Substring(str.Length - length, length)
                : str;
        }

        #endregion
    }
}
