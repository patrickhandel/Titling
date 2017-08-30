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
                    Excel.Range headerRowRange = (Excel.Range)ws.get_Range(sHeaderRangeName, Type.Missing);

                    string header;
                    int column;

                    foreach (Excel.Range cell in headerRowRange.Cells)
                    {
                        int i;
                        header = cell.Value;
                        column = cell.Column;
                        Excel.XlHAlign align;
                        string colType = cell.Offset[-1, 0].Value;
                        
                        switch (colType)
                        {
                            case "TextLong":
                                i = ColumnType.TextLong;
                                align = Excel.XlHAlign.xlHAlignLeft;
                                cell.IndentLevel = 1;
                                break;
                            case "TextMedium":
                                i = ColumnType.TextMedium;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "TextShort":
                                i = ColumnType.TextShort;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "Number":
                                i = ColumnType.Number;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "YesNo":
                                i = ColumnType.YesNo;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "Percent":
                                i = ColumnType.Percent;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "Error":
                                i = ColumnType.Error;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "Date":
                                i = ColumnType.Date;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "Hidden":
                                i = ColumnType.Hidden;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                            default:
                                i = ColumnType.Default;
                                align = Excel.XlHAlign.xlHAlignCenter;
                                break;
                        }
                        cell.HorizontalAlignment = align;
                        cell.EntireColumn.ColumnWidth = i;
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
            var prop = ThisAddIn.wsProps.FirstOrDefault(p => p.Worksheet == name);
            if (prop == null)
                return "";
            return prop.Range + "[#Headers]";
        }
    }

    public struct ColumnType
    {
        public const int TextLong = 40;
        public const int TextMedium = 20;
        public const int TextShort = 15;
        public const int Number	= 9;
        public const int YesNo = 9;
        public const int Percent = 9 ;
        public const int Error = 9;
        public const int Date = 11;
        public const int Default = 15;
        public const int Hidden = 0;
    }
}
