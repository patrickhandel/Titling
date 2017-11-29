using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class WorksheetStandardization
    {
        public static void ExecuteCleanupWorksheet(Excel.Application app)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetSelectedTableHeader(app);
                if (sHeaderRangeName != string.Empty)
                {
                    Range headerRowRange = app.get_Range(sHeaderRangeName, Type.Missing);
                    string header;
                    int column;
                    foreach (Range cell in headerRowRange.Cells)
                    {
                        header = cell.Value;
                        column = cell.Column;
                        string colType = cell.Offset[-1, 0].Value;
                        cell.EntireColumn.ColumnWidth = SSUtils.GetColumnWidth(colType);
                        cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        if (colType == "TextLong")
                        {
                            cell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            cell.IndentLevel = 1;
                        }
                        else
                        {
                            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        }
                    }

                    Worksheet activeWorksheet = app.ActiveSheet;
                    Range r = activeWorksheet.get_Range("A1");
                    r.EntireRow.RowHeight = 40;
                    headerRowRange.EntireRow.RowHeight = 66;
                    headerRowRange.Offset[-1, 0].Font.Size = 9;
                    headerRowRange.Font.Size = 9;
                    headerRowRange.VerticalAlignment = XlVAlign.xlVAlignTop;
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
