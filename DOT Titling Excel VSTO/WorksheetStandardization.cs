using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DOT_Titling_Excel_VSTO
{
    class WorksheetStandardization
    {
        public static void ExecuteCleanupWorksheet(Excel.Worksheet ws)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(ws.Name);
                if (sHeaderRangeName != string.Empty)
                {
                    Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                    string header;
                    int column;
                    foreach (Excel.Range cell in headerRowRange.Cells)
                    {
                        header = cell.Value;
                        column = cell.Column;
                        string colType = cell.Offset[-1, 0].Value;
                        cell.EntireColumn.ColumnWidth = SSUtils.GetColumnWidth(colType);
                        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        if (colType == "TextLong")
                        {
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            cell.IndentLevel = 1;
                        }
                        else
                        {
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }

                    Excel.Range r = ws.get_Range("A1");
                    r.EntireRow.RowHeight = 40;
                    headerRowRange.EntireRow.RowHeight = 66;
                    headerRowRange.Offset[-1, 0].Font.Size = 9;
                    headerRowRange.Font.Size = 9;
                    headerRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
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
