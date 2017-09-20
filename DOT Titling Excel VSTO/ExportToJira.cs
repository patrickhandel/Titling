using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class ExportToJira
    {
        public static void ExecuteSaveSummary()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Worksheet activeWorksheet = app.ActiveSheet;
                Range activeCell = app.ActiveCell;
                if (activeCell != null && activeWorksheet.Name == "Tickets")
                {
                    SSUtils.DoStandardStuff(app);
                    SaveSummary(activeWorksheet, activeCell);
                    SSUtils.DoStandardStuff(app);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void SaveSummary(Worksheet ws, Range activeCell)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Story ID");
                int column = activeCell.Column;
                int row = activeCell.Row;
                string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);
                string newSummary = SSUtils.GetCellValue(ws, row, column);
                JiraUtils.SaveSummary(jiraId, newSummary);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


    }
}
