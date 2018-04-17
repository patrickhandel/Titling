using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class TableViews
    {
        public static void ExecuteViewReleaseNotes_DOT(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Issues"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();

                    ColumnsToShow.Add("Issue Type");
                    ColumnsToShow.Add("Issue ID");
                    SSUtils.SetColumnWidth(ws, "Issue ID", 20);
                    ColumnsToShow.Add("Epic");
                    ColumnsToShow.Add("Summary");
                    SSUtils.SetColumnWidth(ws, "Summary", 150);
                    ColumnsToShow.Add("Epic ID");
                    ColumnsToShow.Add("WIN Release");
                    ColumnsToShow.Add("Epic Release");
                    ColumnsToShow.Add("Agreed Upon Release");
                    ColumnsToShow.Add("DOT Jira ID");
                    ColumnsToShow.Add("Affects Version");

                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                    SSUtils.SortTable(ws, tableRangeName, "WIN Release", Excel.XlSortOrder.xlAscending);
                    SSUtils.SortTable(ws, tableRangeName, "Issue Type", Excel.XlSortOrder.xlDescending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
        public static void ExecuteViewRequirementsErrors(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Issues"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();

                    ColumnsToShow.Add("Issue Type");
                    ColumnsToShow.Add("Issue ID");
                    ColumnsToShow.Add("Link");
                    ColumnsToShow.Add("Summary (Local)");
                    ColumnsToShow.Add("Epic (Local)");
                    ColumnsToShow.Add("Agreed Upon Release");
                    ColumnsToShow.Add("Epic Release");
                    ColumnsToShow.Add("WIN Release");
                    ColumnsToShow.Add("Story Points");
                    ColumnsToShow.Add("Bypass Approval");
                    ColumnsToShow.Add("Sprint");
                    ColumnsToShow.Add("Date Submitted to DOT");
                    ColumnsToShow.Add("Date Approved by DOT");
                    ColumnsToShow.Add("Days Waiting for Approval");
                    ColumnsToShow.Add("ERR Workflow Created");
                    ColumnsToShow.Add("ERR Workflow Written");
                    ColumnsToShow.Add("ERR Workflow Groomed");
                    ColumnsToShow.Add("ERR Workflow Submitted");
                    ColumnsToShow.Add("ERR Workflow Ready");
                    ColumnsToShow.Add("ERR Workflow Approved Not Groomed");
                    ColumnsToShow.Add("ERR Workflow Bug Bucket");
                    ColumnsToShow.Add("Has Workflow Issue");

                    SSUtils.FilterTable(ws, tableRangeName, "Has Workflow Issue", "x");
                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                    SSUtils.SortTable(ws, tableRangeName, "Sprint", Excel.XlSortOrder.xlAscending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteViewEpicsEstimateActual(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Epics"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();

                    ColumnsToShow.Add("Epic");
                    ColumnsToShow.Add("R");
                    ColumnsToShow.Add("Estimate 3");
                    ColumnsToShow.Add("Actual");
                    ColumnsToShow.Add("Actual vs Estimate");

                    SSUtils.FilterTable(ws, tableRangeName, "Release Number", "<8");
                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                    SSUtils.SortTable(ws, tableRangeName, "Priority", Excel.XlSortOrder.xlAscending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteViewBlockedIssues_DOT(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Issues"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();

                    ColumnsToShow.Add("Issue Type");
                    ColumnsToShow.Add("Issue ID");
                    ColumnsToShow.Add("Link");
                    ColumnsToShow.Add("Epic (Local)");
                    ColumnsToShow.Add("Summary (Local)");
                    ColumnsToShow.Add("Story Points");
                    ColumnsToShow.Add("WIN Release");
                    ColumnsToShow.Add("Status");
                    ColumnsToShow.Add("Status (Last Changed)");
                    ColumnsToShow.Add("Days in Same Status");
                    ColumnsToShow.Add("Assignee");
                    ColumnsToShow.Add("ERR Story Not Moving or Blocked");
                    ColumnsToShow.Add("Reason Blocked or Delayed");
                    ColumnsToShow.Add("ERR Need Reason for Blocker");

                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                    SSUtils.FilterTable(ws, tableRangeName, "ERR Story Not Moving or Blocked", "x");
                    SSUtils.SortTable(ws, tableRangeName, "Assignee", Excel.XlSortOrder.xlAscending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteViewReleasePlan_DOT(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Releases"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();
                    ColumnsToShow.Add("R");
                    ColumnsToShow.Add("Release");
                    ColumnsToShow.Add("Status");
                    ColumnsToShow.Add("FRS Delivery Date");
                    ColumnsToShow.Add("Mid/Long");
                    ColumnsToShow.Add("From (Date)");
                    ColumnsToShow.Add("To (Date)");
                    ColumnsToShow.Add("UAT From (Date)");
                    ColumnsToShow.Add("UAT To (Date)");
                    ColumnsToShow.Add("Vendor Release");
                    ColumnsToShow.Add("Deliver to Vendors To (Reported)");
                    ColumnsToShow.Add("Deliver to Vendors To (Actual)");

                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                    SSUtils.SortTable(ws, tableRangeName, "Release Number", Excel.XlSortOrder.xlAscending);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void ExecuteViewRequirementsStatus_DOT(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.Worksheets["Issues"];
                ws.Activate();

                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    int headerRow = headerRowRange.Row;

                    List<string> ColumnsToShow = new List<string>();
                    ColumnsToShow.Add("Issue Type");
                    ColumnsToShow.Add("Issue ID");
                    ColumnsToShow.Add("Link");
                    ColumnsToShow.Add("Epic (Local)");
                    ColumnsToShow.Add("Story Points");
                    ColumnsToShow.Add("WIN Release");
                    ColumnsToShow.Add("Sprint Number (Local)");
                    ColumnsToShow.Add("Summary");
                    ColumnsToShow.Add("Sprint");
                    ColumnsToShow.Add("ID");
                    ColumnsToShow.Add("Sprint Number");
                    ColumnsToShow.Add("Date Submitted to DOT");
                    ColumnsToShow.Add("Date Approved by DOT");
                    ColumnsToShow.Add("Bypass Approval");
                    ColumnsToShow.Add("Requirements Gathering Notes");

                    SSUtils.HideTableColumns(headerRowRange, ColumnsToShow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
