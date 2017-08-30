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
        public static void ExecuteCleanup()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Workbook wb = app.ActiveWorkbook;
                Excel.Worksheet ws = wb.ActiveSheet;

                string sHeaderRangeName = GetHeaderRangeName(ws.Name);
                if (sHeaderRangeName != "")
                {
                    app.ScreenUpdating = false;
                    Excel.Range headerRowRange = (Excel.Range)ws.get_Range(sHeaderRangeName, Type.Missing);
                    string header;
                    int column;
                    foreach (Excel.Range cell in headerRowRange.Cells)
                    {
                        header = cell.Value;
                        column = cell.Column;
                        string colType = cell.Offset[-1, 0].Value;
                        cell.EntireColumn.ColumnWidth = GetColumnWidth(colType);
                        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        if (colType == "TextLong")
                        {
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        else
                        {
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            cell.IndentLevel = 1;
                        }
                    }

                    Excel.Range r = ws.get_Range("A1");
                    r.EntireRow.RowHeight = 40;
                    headerRowRange.EntireRow.RowHeight = 66;
                    headerRowRange.Offset[-1, 0].Font.Size = 10;
                    headerRowRange.Font.Size = 10;
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;
                    app.ScreenUpdating = true;
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static string GetHeaderRangeName(string name)
        {
            List<WorksheetProperties> wsProps = WorksheetPropertiesManager.GetWorksheetProperties();
            var prop = wsProps.FirstOrDefault(p => p.Worksheet == name);
                if (prop == null)
                    return "";
                return prop.Range + "[#Headers]";
        }

        private static int GetColumnWidth(string name)
        {
            List<ColumnTypes> wsColumnTypes = WorksheetPropertiesManager.GetColumnTypes();
            var prop = wsColumnTypes.FirstOrDefault(p => p.Name == name);
            if (prop == null)
                return 15;
            return (int)prop.Width;
        }
    }
}
