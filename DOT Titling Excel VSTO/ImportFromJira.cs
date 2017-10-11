using System;
using System.Windows.Forms;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Atlassian.Jira;
using System.Collections.Generic;

namespace DOT_Titling_Excel_VSTO
{
    class ImportFromJira
    {
        public static void ExecuteUpdateSelectedTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;

                if (activeCell != null && ((activeWorksheet.Name == "Tickets") || (activeWorksheet.Name == "DOT Releases")))
                {
                    UpdateSelectedTickets(app, activeWorksheet, selection);
                    WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteAddNewTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Tickets") || (activeWorksheet.Name == "DOT Releases"))
                {
                    AddNewTickets(app, activeWorksheet);
                    WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewEpics(Excel.Application app, Worksheet ws)
        {
            try
            {
                var issues = JiraUtils.GetAllIssues("Epics").Result;
                string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                int column = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(wsRangeName);

                List<string> listOfTicketIDs = new List<string>();
                Range ticketIDColumnRange = ws.get_Range(wsRangeName + "[Epic ID]", Type.Missing);
                foreach (Range cell in ticketIDColumnRange.Cells)
                {
                    listOfTicketIDs.Add(cell.Value);
                }
                foreach (var ticketID in listOfTicketIDs)
                {
                    issues.Remove(issues.FirstOrDefault(x => x.Key.Value == ticketID.ToString()));
                }

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                foreach (var issue in issues)
                {
                    Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                    int footerRow = footerRangeRange.Row;
                    Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                    rToInsert.Insert();
                    UpdateValues(ws, jiraFields, footerRow, issue, false);
                    SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value);
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), issue.Summary);
                    SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                }
                MessageBox.Show(issues.Count() + " Tickets Added.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void AddNewTickets(Excel.Application app, Worksheet ws)
        {
            try
            {
                var issues = JiraUtils.GetAllIssues().Result;
                string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                int column = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(wsRangeName);

                List<string> listOfTicketIDs = new List<string>();
                Range ticketIDColumnRange = ws.get_Range(wsRangeName + "[Ticket ID]", Type.Missing);
                foreach (Range cell in ticketIDColumnRange.Cells)
                {
                    listOfTicketIDs.Add(cell.Value);
                }
                foreach (var ticketID in listOfTicketIDs)
                {
                    issues.Remove(issues.FirstOrDefault(x => x.Key.Value == ticketID.ToString()));
                }

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                foreach (var issue in issues)
                {
                    Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                    int footerRow = footerRangeRange.Row;
                    Range rToInsert = ws.get_Range(String.Format("{0}:{0}", footerRow), Type.Missing);
                    rToInsert.Insert();
                    UpdateValues(ws, jiraFields, footerRow, issue, false);
                    SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value);
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Summary"), issue.Summary);
                    //TO DO FIX
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Release")));

                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    int jiraEpicColumn = SSUtils.GetColumnFromHeader(ws, "Jira Epic");
                    string newEpic = SSUtils.GetCellValue(ws, footerRow, jiraEpicColumn);
                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), newEpic);
                    app.Calculation = XlCalculation.xlCalculationManual;

                    SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Hufflepuff Sprint"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Hufflepuff Sprint")));
                    SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                }
                MessageBox.Show(issues.Count() + " Tickets Added.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteUpdateAllTickets()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Tickets") || (activeWorksheet.Name == "DOT Releases"))
                {
                    UpdateAllTickets(app, activeWorksheet);
                    AddNewTickets(app, activeWorksheet);
                    string dt = DateTime.Now.ToString("mm/dd/yyyy");
                    if (activeWorksheet.Name == "Tickets")
                    {
                        string val = "Tickets (Updated on " + dt + ")";
                        SSUtils.SetCellValue(activeWorksheet, 1, 1, val);
                    }
                    if (activeWorksheet.Name == "DOT Releases")
                    {
                        string val = "DOT Releases (Updated on " + dt + ")";
                        SSUtils.SetCellValue(activeWorksheet, 1, 1, val);
                    }
                    WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteUpdateEpics()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Epics"))
                {
                    UpdateEpics(app, activeWorksheet);
                    AddNewEpics(app, activeWorksheet);
                    WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


        public static void ExecuteUpateTicketBeforeMailMerge(string jiraId)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ws = app.Sheets["Tickets"];
                UpdateTicketBeforeMailMerge(app, ws, jiraId);
                WorksheetStandardization.ExecuteCleanupWorksheet(ws);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateTicketBeforeMailMerge(Excel.Application app, Worksheet ws, string jiraId)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(rangeName);
                var issue  = JiraUtils.GetIssue(jiraId).Result;
                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                int row = SSUtils.FindTextInColumn(ws, rangeName + "[Ticket ID]", jiraId);
                bool notFound = issue == null;
                UpdateValues(ws, jiraFields, row, issue, notFound);
                SSUtils.SetStandardRowHeight(ws, row, row);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateEpics(Excel.Application app, Worksheet ws)
        {
            try
            {
                var epics = JiraUtils.GetAllIssues("Epics").Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields("EpicData");

                int cnt = epics.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string jiraID = SSUtils.GetCellValue(ws, currentRow, jiraIDCol);
                    var issue = epics.FirstOrDefault(i => i.Key == jiraID);
                    bool notFound = issue == null;
                    UpdateValues(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateAllTickets(Excel.Application app, Worksheet ws)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
            try
            {
                var issues = JiraUtils.GetAllIssues().Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields("TicketData");

                int cnt = issues.Count();

                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(ws.Name);
                Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                for (int currentRow = headerRow + 1; currentRow < footerRow; currentRow++)
                {
                    string jiraID = SSUtils.GetCellValue(ws, currentRow, jiraIDCol);
                    var issue = issues.FirstOrDefault(i => i.Key == jiraID);
                    bool notFound = issue == null;
                    UpdateValues(ws, jiraFields, currentRow, issue, notFound);
                }
                SSUtils.SetStandardRowHeight(ws, headerRow + 1, footerRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateSelectedTickets(Excel.Application app, Worksheet ws, Range selection)
        {
            //// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home

            var issues = JiraUtils.GetAllIssues().Result;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields("TicketData");

            string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
            Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            int cnt = selection.Rows.Count;

            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                    string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol).Trim();
                    if (jiraId.Length > 10 && jiraId.Substring(0, 10) == "DOTTITLNG-")
                    {
                        var issue = issues.FirstOrDefault(p => p.Key.Value == jiraId);
                        bool notFound = issue == null;
                        UpdateValues(ws, jiraFields, row, issue, notFound);
                        SSUtils.SetStandardRowHeight(ws, row, row);
                    }
                }
            }
        }

        private static void UpdateValues(Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Issue issue, bool notFound)
        {
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
                        int ticketTypeCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Ticket Type");
                        SSUtils.SetCellValue(activeWorksheet, row, ticketTypeCol, valueToSave);
                    }
                    SSUtils.SetCellValue(activeWorksheet, row, column, valueToSave);
                }
                else
                {
                    if (type == "Standard")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraUtils.ExtractStandardValue(issue, item));
                    if (type == "Custom")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraUtils.ExtractCustomValue(issue, item));
                    if (type == "Function")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraUtils.ExtractValueBasedOnFunction(issue, item));
                }
                if (type == "Formula")
                    SSUtils.SetCellFormula(activeWorksheet, row, column, formula);
            }
        }
    }
}
