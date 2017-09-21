using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class ExportToJira
    {
        public static void ExecuteSaveTicket()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Worksheet activeWorksheet = app.ActiveSheet;
                Range activeCell = app.ActiveCell;
                if (activeCell != null && activeWorksheet.Name == "Tickets")
                {
                    SaveTicket(activeWorksheet, activeCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void SaveTicket(Worksheet ws, Range activeCell)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                
                int column = activeCell.Column;
                int row = activeCell.Row;
                string fieldToSave = SSUtils.GetCellValue(ws, headerRowRange.Row, column);
                string newValue = SSUtils.GetCellValue(ws, row, column).Trim();

                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);

                int typeCol = SSUtils.GetColumnFromHeader(ws, "Issue Type");
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
                            JiraUtils.SaveSummary(jiraId, newValue);
                            break;
                        case "Jira Status":
                            JiraUtils.SaveStatus(jiraId, newValue);
                            break;
                        case "Date Submitted to DOT":
                            if (type == "Story")
                            {
                                newValue = newValue.Trim();
                                if (newValue != string.Empty)
                                {
                                    if (CheckDate(newValue) == false)
                                    {
                                        MessageBox.Show(fieldToSave + " is not a valid date.");
                                        break;
                                    }
                                    DateTime dt = DateTime.Parse(newValue);
                                    newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                }
                                JiraUtils.SaveCustomField(jiraId, fieldToSave, newValue);
                            }
                            MessageBox.Show(fieldToSave + " can't be updated because it is not a story.");
                            break;
                        case "Date Approved by DOT":
                            if (type == "Story")
                            {
                                newValue = newValue.Trim();
                                if (newValue != string.Empty)
                                {
                                    if (CheckDate(newValue) == false)
                                    {
                                        MessageBox.Show(fieldToSave + " is not a valid date.");
                                        break;
                                    }
                                    DateTime dt = DateTime.Parse(newValue);
                                    newValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz").Remove(26, 1);
                                }
                                JiraUtils.SaveCustomField(jiraId, fieldToSave, newValue);
                            }
                            MessageBox.Show(fieldToSave + " can't be updated because it is not a story.");
                            break;
                        case "Story - As A":
                            if (type == "Story")
                            {
                                JiraUtils.SaveCustomField(jiraId, "Story: As a(n)", newValue);
                            }
                            else
                            {
                                MessageBox.Show(fieldToSave + " can't be updated because it is not a story.");
                            }
                            break;
                        case "Story - Id Like":
                            if (type == "Story")
                            {
                                JiraUtils.SaveCustomField(jiraId, "Story: I'd like to be able to", newValue);
                            }
                            else
                            {
                                MessageBox.Show(fieldToSave + " can't be updated because it is not a story.");
                            }
                            break;
                        case "Story - So That":
                            if (type == "Story")
                            {
                                JiraUtils.SaveCustomField(jiraId, "Story: So that I can", newValue);
                            }
                            else
                            {
                                MessageBox.Show(fieldToSave + " can't be updated because it is not a story.");
                            }
                            break;
                        case "Points":
                            JiraUtils.SaveCustomField(jiraId, "Story Points", newValue);
                            break;
                        case "DOT Jira ID":
                            if (type == "Software Bug")
                            {
                                JiraUtils.SaveCustomField(jiraId, fieldToSave, newValue);
                            }
                            else
                            {
                                MessageBox.Show(fieldToSave + " can't be updated because it is not a Software Bug.");
                            }
                            break;
                        case "Jira Release":
                            JiraUtils.SaveRelease(jiraId, newValue);
                            break;
                        case "Jira Epic ID":
                            JiraUtils.SaveCustomField(jiraId, "Epic Link", newValue);
                            break;
                        default:
                            //DO NOT UPDATE THE FOLLOWING:
                            //Issue Type
                            //Jira Fix Release
                            //Jira Hufflepuff Sprint
                            //Jira Epic
                            MessageBox.Show(fieldToSave + " can't be updated.");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
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
}
