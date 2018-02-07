using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class ExportToJira
    {
        public static bool SaveSelectedIssues(Excel.Application app)
        {
            try
            {
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                var selection = app.Selection;
                if (activeCell != null && (activeWorksheet.Name == "Issues" || activeWorksheet.Name == "DOT Releases"))
                {
                    return SaveSelectedIssueValues(activeWorksheet, selection);
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

        private static bool SaveSelectedEpic(Excel.Worksheet ws, Excel.Range activeCell)
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
                        JiraIssue.SaveSummary(issueID, newValue, multiple);
                        break;
                    case "Status":
                        JiraIssue.SaveStatus(issueID, newValue, multiple);
                        break;
                    case "Story Points":
                        JiraIssue.SaveCustomField(issueID, "Story Points", newValue, multiple);
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

        private static bool SaveSelectedIssueValues(Excel.Worksheet ws, Excel.Range selection)
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
                                    JiraIssue.SaveSummary(issueID, newValue, multiple);
                                    break;
                                case "Status":
                                    JiraIssue.SaveStatus(issueID, newValue, multiple);
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
                                        JiraIssue.SaveCustomField(issueID, fieldToSave, newValue, multiple);
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
                                        JiraIssue.SaveCustomField(issueID, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - As A":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(issueID, "Story: As a(n)", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - Id Like":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(issueID, "Story: I'd like to be able to", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story - So That":
                                    if (type == "Story")
                                    {
                                        JiraIssue.SaveCustomField(issueID, "Story: So that I can", newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a story. (" + row + ")");
                                    }
                                    break;
                                case "Story Points":
                                    JiraIssue.SaveCustomField(issueID, "Story Points", newValue, multiple);
                                    break;
                                case "DOT Jira ID":
                                    if (type == "Software Bug")
                                    {
                                        JiraIssue.SaveCustomField(issueID, fieldToSave, newValue, multiple);
                                    }
                                    else
                                    {
                                        MessageBox.Show(fieldToSave + " can't be updated because it is not a Software Bug. (" + row + ")");
                                    }
                                    break;
                                case "Release":
                                    JiraIssue.SaveRelease(issueID, newValue, multiple);
                                    break;
                                case "Labels":
                                    JiraIssue.SaveLabels(issueID, newValue, multiple);
                                    break;
                                case "Epic Link":
                                    JiraIssue.SaveCustomField(issueID, "Epic Link", newValue, multiple);
                                    break;
                                case "SWAG":
                                    JiraIssue.SaveCustomField(issueID, "SWAG", newValue, multiple);
                                    break;
                                case "Reason Blocked or Delayed":
                                    JiraIssue.SaveCustomField(issueID, "Reason Blocked or Delayed", newValue, multiple);
                                    break;
                                //case "Sprint":
                                //    JiraUtils.SaveCustomField(issueID, "Sprint", newValue);
                                //    break;
                                default:
                                    //DO NOT UPDATE THE FOLLOWING:
                                    //Issue Type
                                    //Sprint Number
                                    //Epic
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
