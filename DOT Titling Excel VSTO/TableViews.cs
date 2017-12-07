using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace DOT_Titling_Excel_VSTO
{
    class TableViews
    {
        public static void ExecuteViewBlockedTickets(Excel.Application app)
        {
            try
            {
                Worksheet ws = app.Worksheets["Tickets"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();

                    ColumnsToShow.Add("Ticket Type");
                    ColumnsToShow.Add("Ticket ID");
                    ColumnsToShow.Add("Link");
                    ColumnsToShow.Add("Epic");
                    ColumnsToShow.Add("Summary");
                    ColumnsToShow.Add("Points");
                    ColumnsToShow.Add("WIN Release");
                    ColumnsToShow.Add("Jira Status");
                    ColumnsToShow.Add("Jira Status (Last Changed)");
                    ColumnsToShow.Add("Days in Same Status");
                    ColumnsToShow.Add("Assignee");
                    ColumnsToShow.Add("ERR Story Not Moving or Blocked");
                    ColumnsToShow.Add("Reason Blocked or Delayed");

                    HideTableColumns(headerRowRange, ColumnsToShow);

                    //Range tableRange = app.get_Range(tableRangeName, Type.Missing);
                    //SSUtils.SortAndFilterTable(ws, tableRange, "ERR Story Not Moving or Blocked", XlSortOrder.xlDescending, "ERR Story Not Moving or Blocked", "x");
                    //SSUtils.SortTable(ws, tableRangeName, "ERR Story Not Moving or Blocked", XlSortOrder.xlDescending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteViewReleasePlan(Excel.Application app)
        {
            try
            {
                Worksheet ws = app.Worksheets["Releases"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();
                    ColumnsToShow.Add("R");
                    ColumnsToShow.Add("Mid/Long");
                    ColumnsToShow.Add("From (Date)");
                    ColumnsToShow.Add("To (Date)");
                    ColumnsToShow.Add("UAT From (Date)");
                    ColumnsToShow.Add("UAT To (Date)");
                    ColumnsToShow.Add("Deliver to Vendors");

                    HideTableColumns(headerRowRange, ColumnsToShow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void HideTableColumns(Range headerRowRange, List<string> ColumnsToShow)
        {
            // Format each cell in the table header row
            foreach (Range cell in headerRowRange.Cells)
            {
                int column = cell.Column;
                string columnHeader = cell.Value;

                var item = ColumnsToShow.Find(x => x == columnHeader);
                if (item == null)
                {
                    cell.EntireColumn.ColumnWidth = 0;
                }
            }
        }
    }
}
