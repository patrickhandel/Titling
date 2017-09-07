using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Atlassian.Jira;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace DOT_Titling_Excel_VSTO
{
    class Import
    {
        //// Atlassian.NET SDK
        //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

        public static void JiraTickets()
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Workbook wb = app.ActiveWorkbook;
            Worksheet ws = wb.ActiveSheet;
            try
            {
                if (ws.Name == "Jira Tickets")
                {
                    app.ScreenUpdating = false;

                    var jira = Jira.CreateRestClient("https://wiportal.atlassian.net", "patrick.handel@egov.com", "viPer47,,");
                    jira.MaxIssuesPerRequest = 1000;
                    var issues = from i in jira.Issues.Queryable
                                 where i.Project == "DOTTITLNG" &&
                                    (i.Type == "Story" || i.Type == "Software Bug") &&
                                    i.Summary != "DELETE"
                                 orderby i.Created
                                 select i;
                    int cnt = issues.Count();

                    string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                    string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);

                    Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                    Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);

                    int headerRow = headerRowRange.Row;
                    int footerRow = footerRangeRange.Row;

                    Range rToInsert = ws.get_Range(String.Format("{0}:{1}", footerRow, footerRow + cnt - 1), Type.Missing);
                    Range rToDelete = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);

                    rToInsert.Insert();
                    rToDelete.Delete();

                    int row = headerRow + 1;
                    foreach (var issue in issues)
                    {
                        //Populate the row
                        foreach (Range cell in headerRowRange.Cells)
                        {
                            string header = cell.Value;
                            int column = cell.Column;
                            SSUtils.SetCellValue(ws, row, column, GetIssueValueForCell(issue, header));
                        }
                        row++;
                    }

                    SetStandardRowHeight(ws, headerRow, footerRow);
                    SetColumnsWithFormulas(ws, headerRow, footerRow);
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                app.ScreenUpdating = true;
            }
        }

        private static void SetStandardRowHeight(Worksheet ws, int headerRow, int footerRow)
        {
            Range allRows = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
            allRows.EntireRow.RowHeight = 15;
        }

        private static void SetColumnsWithFormulas(Worksheet ws, int headerRow, int footerRow)
        {
            Range caclCol;
            string formula;
            //Epic
            caclCol = ws.get_Range("JiraTicketData[Epic]", Type.Missing);
            formula = "=IFERROR(INDEX(JiraEpicData[Summary],MATCH([@[Epic ID]],JiraEpicData[Jira Epic ID],0)),\"\")";
            caclCol.Formula = string.Format(formula);

            //ERR Ticket Not Found
            caclCol = ws.get_Range("JiraTicketData[ERR Ticket Not Found]", Type.Missing);
            formula = "=IF(ISERROR(INDEX(TicketData[Story ID],MATCH([@[Story ID]],TicketData[Story ID],0))),\"x\",\"\")";
            caclCol.Formula = string.Format(formula);

            //ERR No Epic
            caclCol = ws.get_Range("JiraTicketData[ERR No Epic]", Type.Missing);
            formula = "=IF([@Epic]=\"\",\"x\",\"\")";
            caclCol.Formula = string.Format(formula);

            //ERR Points but To Do
            caclCol = ws.get_Range("JiraTicketData[ERR Points but To Do]", Type.Missing);
            formula = "=IF(AND([@[Issue Type]]=\"Story\",[@Status]=\"To Do\",[@[Story Points]]<>0),\"x\",\"\")";
            caclCol.Formula = string.Format(formula);

            //ERR Done No Sprint
            caclCol = ws.get_Range("JiraTicketData[ERR Done No Sprint]", Type.Missing);
            formula = "=IF([@Status]=\"Done\",IF([@Sprint]=\"\",\"x\",\"\"),\"\")";
            caclCol.Formula = string.Format(formula);

            //ERR Bug Not Categorized
            caclCol = ws.get_Range("JiraTicketData[ERR Bug Not Categorized]", Type.Missing);
            formula = "=IF(AND([@[Issue Type]]=\"Software Bug\",[@[DOT Jira ID]]=\"\"),\"x\",\"\")";
            caclCol.Formula = string.Format(formula);
        }

        private static string GetIssueValueForCell(Issue issue, string header)
        {
            string val = string.Empty;
            switch (header)
            {
                case "Issue Type":
                    val = issue.Type.Name;
                    break;
                case "Story ID":
                    val = issue.Key.Value;
                    break;
                case "Epic ID":
                    val = SSUtils.GetCustomValue(issue, "Epic Link");
                    break;
                case "Summary":
                    val = issue.Summary;
                    break;
                case "Status":
                    val = issue.Status.Name;
                    break;
                case "Story Points":
                    val = SSUtils.GetCustomValue(issue, "Story Points");
                    break;
                case "Date Submitted to DOT":
                    val = SSUtils.GetCustomValue(issue, "Date Submitted to DOT"); ;
                    break;
                case "Date Approved by DOT":
                    val = SSUtils.GetCustomValue(issue, "Date Approved by DOT");
                    break;
                case "DOT Jira ID":
                    val = SSUtils.GetCustomValue(issue, "DOT Jira ID");
                    break;
                case "Description":
                    val = issue.Description;
                    break;
                case "Story - As An":
                    val = SSUtils.GetCustomValue(issue, "Story: As a(n)"); ;
                    break;
                case "Story - Id Like":
                    val = SSUtils.GetCustomValue(issue, "Story: I'd like to be able to");
                    break;
                case "Story - So That":
                    val = SSUtils.GetCustomValue(issue, "Story: So that I can");
                    break;
                case "Sprint":
                    val = ExtractSprintNumber(issue);
                    break;
                case "Release":
                    val = ExtractRelease(issue);
                    break;
                case "Fix Release":
                    val = ExtractFixRelease(issue);
                    break;
                default:
                    break;
            }

            return val;
        }

        private static string ExtractRelease(Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.AffectsVersions)
            {
                val = issue.AffectsVersions[c].Name;
                c++;
            }
            return val;
        }

        private static string ExtractFixRelease(Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.FixVersions)
            {
                val = issue.FixVersions[c].Name;
                c++;
            }
            return val;
        }

        private static string ExtractSprintNumber(Issue issue)
        {
            string val = SSUtils.GetCustomValue(issue, "Sprint");
            if (val != string.Empty)
            {
                val = string.Empty;
                foreach (var value in issue.CustomFields["Sprint"].Values)
                    val = value;
                val = val.Replace("DOT", "");
                val = val.Replace("Backlog", "");
                val = val.Replace("Hufflepuff", "");
                val = val.Replace("Sprint", "");
                val = val.Replace("Ready", "");
                val = val.Replace("Other", "");
                val = val.Replace("-", "");
                val = val.Replace(" ", "");
                for (int rev = 1; rev <= 12; rev++)
                    val = val.Replace("R" + rev.ToString(), "");
            }
            return val;
        }
    }
}
