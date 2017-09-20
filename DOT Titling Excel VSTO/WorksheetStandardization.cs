using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;

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
                    SSUtils.DoStandardStuff(app);
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
                    headerRowRange.Offset[-1, 0].Font.Size = 10;
                    headerRowRange.Font.Size = 10;
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;
                    SSUtils.DoStandardStuff(app);
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
