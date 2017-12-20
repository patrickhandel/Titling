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
        public static void ExecuteUpdateAllTickets(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Tickets") || (activeWorksheet.Name == "DOT Releases"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateAllTickets(app, activeWorksheet);
                        AddNewTickets(app, activeWorksheet);
                        string dt = DateTime.Now.ToString("MM/dd/yyyy");
                        if (activeWorksheet.Name == "Tickets")
                        {
                            string val = "Tickets (Updated on " + dt + ")";
                            SSUtils.SetCellValue(activeWorksheet, 1, 1, val, "Ticket Updated On");
                        }
                        if (activeWorksheet.Name == "DOT Releases")
                        {
                            string val = "DOT Releases (Updated on " + dt + ")";
                            SSUtils.SetCellValue(activeWorksheet, 1, 1, val, "DOT Releases Updated On");
                        }
                        TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteUpdateSelectedTickets(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;
                string table = SSUtils.GetSelectedTable(app);
                if (activeCell != null && ((table == "TicketData") || (table == "DOTReleaseData")))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateSelectedTickets(app, activeWorksheet, selection);
                        //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteAddNewTickets(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Tickets") || (activeWorksheet.Name == "DOT Releases"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        AddNewTickets(app, activeWorksheet);
                        //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        public static void ExecuteUpateTicketBeforeMailMerge(string jiraId)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ws = app.Sheets["Tickets"];
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    UpdateTicketBeforeMailMerge(app, ws, jiraId);
                    //TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        private static void AddNewTickets(Excel.Application app, Worksheet ws)
        {
            try
            {
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = JiraIssue.GetAllIssues().Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int column = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

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

                        //Ticket ID (2)
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "Ticket ID");

                        //Summary (5)
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Summary"), issue.Summary, "Summary");

                        //TO DO FIX
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Release"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Release")), "Release");

                        //Epic
                        app.Calculation = XlCalculation.xlCalculationAutomatic;
                        int jiraEpicColumn = SSUtils.GetColumnFromHeader(ws, "Jira Epic");
                        string newEpic = SSUtils.GetCellValue(ws, footerRow, jiraEpicColumn);
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), newEpic, "Epic");
                        app.Calculation = XlCalculation.xlCalculationManual;

                        //Hufflepuff Sprint
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Hufflepuff Sprint"), SSUtils.GetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Jira Hufflepuff Sprint")), "Hufflepuff Sprint");
                        SSUtils.SetStandardRowHeight(ws, footerRow, footerRow);
                    }
                    MessageBox.Show(issues.Count() + " Tickets Added.");

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

        public static void ExecuteUpdateEpics(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if ((activeWorksheet.Name == "Epics"))
                {
                    string missingColumns = MissingColumns(activeWorksheet);
                    if (missingColumns == string.Empty)
                    {
                        UpdateEpics(app, activeWorksheet);
                        AddNewEpics(app, activeWorksheet);
                        TableStandardization.ExecuteCleanupTable(app, TableStandardization.StandardizationType.Light);
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

        private static void AddNewEpics(Excel.Application app, Worksheet ws)
        {
            try
            {
                string missingColumns = MissingColumns(ws);
                if (missingColumns == string.Empty)
                {
                    var issues = JiraIssue.GetAllIssues("Epics").Result;
                    string wsRangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                    int column = SSUtils.GetColumnFromHeader(ws, "Epic ID");
                    var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

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
                        SSUtils.SetCellValue(ws, footerRow, column, issue.Key.Value, "issue.Key.Value");
                        SSUtils.SetCellValue(ws, footerRow, SSUtils.GetColumnFromHeader(ws, "Epic"), issue.Summary, "Summary");
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

        private static void UpdateTicketBeforeMailMerge(Excel.Application app, Worksheet ws, string jiraId)
        {
            try
            {
                string rangeName = SSUtils.GetWorksheetRangeName(ws.Name);
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
                var issue  = JiraIssue.GetIssue(jiraId).Result;
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
                var epics = JiraIssue.GetAllIssues("Epics").Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

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
                var issues = JiraIssue.GetAllIssues().Result;
                var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);





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

            var issues = JiraIssue.GetAllIssues().Result;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);

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
            //Get the current status
            int statusColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "Jira Status");            
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
                        int ticketTypeCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Ticket Type");
                        if (ticketTypeCol != 0)
                            SSUtils.SetCellValue(activeWorksheet, row, ticketTypeCol, valueToSave, columnHeader);
                    }
                    SSUtils.SetCellValue(activeWorksheet, row, column, valueToSave, columnHeader);
                }
                else
                {
                    if (type == "Standard")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractStandardValue(issue, item), columnHeader);
                    if (type == "Custom")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractCustomValue(issue, item), columnHeader);
                    if (type == "Function")
                        SSUtils.SetCellValue(activeWorksheet, row, column, JiraIssue.ExtractValueBasedOnFunction(issue, item), columnHeader);
                }
                if (type == "Formula")
                    SSUtils.SetCellFormula(activeWorksheet, row, column, formula);
            }

            if (statusColumn != 0)
            {
                newStatus = SSUtils.GetCellValue(activeWorksheet, row, statusColumn);
                int statusLastChangedColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "Jira Status (Last Changed)");
                if (statusLastChangedColumn != 0)
                {
                    string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");
                    int sprintColumn = SSUtils.GetColumnFromHeader(activeWorksheet, "DOT Sprint");
                    string sprint = SSUtils.GetCellValue(activeWorksheet, row, sprintColumn);

                    if (sprint != currentSprint)
                    {
                        SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, string.Empty, "Jira Status (Last Changed)");
                    }
                    else
                    {
                        if (newStatus == "Done" || newStatus == "Ready for Development" || newStatus == "")
                        {
                            SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, string.Empty, "Jira Status (Last Changed)");
                        }
                        else
                        if (newStatus != previousStatus)
                        {
                            SSUtils.SetCellValue(activeWorksheet, row, statusLastChangedColumn, DateTime.Now.ToString("MM/dd/yyyy"), "Jira Status (Last Changed)");
                        }
                    }
                }
            }
        }

        public static string MissingColumns(Worksheet ws)
        {
            string missingFields = string.Empty;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                if (SSUtils.GetColumnFromHeader(ws, columnHeader) == 0)
                    missingFields = missingFields + ' ' + columnHeader;
            }
            return missingFields.Trim();
        }

        private static void UpdateSprintProgress(Worksheet activeWorksheet, List<JiraFields> jiraFields, int row, Issue issue, bool notFound)
        {
            // Get the Current Date
            string dt = DateTime.Now.ToString("MM/dd/yyyy");

            // Get the Current Sprint
            string currentSprint = SSUtils.GetCellValueFromNamedRange("CurrentSprintToUse");

            //Get List of Tickets
            List<Ticket> tickets = Lists.GetListOfTickets(activeWorksheet);
            tickets = tickets.FindAll(r => r.Sprint == currentSprint && r.Type == "Story");
            
            //TO DO
            //Add each ticket to the bottom of the Table
        }
    }
}
