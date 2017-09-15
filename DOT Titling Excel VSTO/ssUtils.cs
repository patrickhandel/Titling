using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class SSUtils
    {
        public static int GetColumnFromHeader(Worksheet ws, string columnText)
        {
            try
            {
                string sHeaderRangeName = GetHeaderRangeName(ws.Name);
                Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                foreach (Range cell in headerRowRange.Cells)
                {
                    if (cell.Value == columnText)
                        return cell.Column;
                }
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return 0;
            }
        }

        public static string GetWorksheetRangeName(string name)
        {
            List<WorksheetProperties> wsProps = WorksheetPropertiesManager.GetWorksheetProperties();
            var prop = wsProps.FirstOrDefault(p => p.Worksheet == name);
            if (prop == null)
                return string.Empty;
            return prop.Range;
        }

        public static string GetHeaderRangeName(string name)
        {
            List<WorksheetProperties> wsProps = WorksheetPropertiesManager.GetWorksheetProperties();
            var prop = wsProps.FirstOrDefault(p => p.Worksheet == name);
            if (prop == null)
                return string.Empty;
            return prop.Range + "[#Headers]";
        }

        public static string GetFooterRangeName(string name)
        {
            List<WorksheetProperties> wsProps = WorksheetPropertiesManager.GetWorksheetProperties();
            var prop = wsProps.FirstOrDefault(p => p.Worksheet == name);
            if (prop == null)
                return string.Empty;
            return prop.Range + "[#Totals]";
        }

        public static int GetColumnWidth(string name)
        {
            List<ColumnTypes> wsColumnTypes = WorksheetPropertiesManager.GetColumnTypes();
            var prop = wsColumnTypes.FirstOrDefault(p => p.Name == name);
            if (prop == null)
                return 15;
            return prop.Width;
        }

        public static string GetCellValue(Excel.Worksheet sheet, int row, int column)
        {
            var result = string.Empty;
            if (sheet != null)
            {
                Range rng = sheet.Cells[row, column] as Range;

                if (rng != null)
                    result = (string)rng.Text;
            }
            return (result + " ").Trim();
        }

        public static void SetCellValue(Worksheet sheet, int row, int column, string val)
        {
            if (sheet != null)
            {
                Excel.Range rng = sheet.Cells[row, column] as Excel.Range;
                if (rng != null)
                    rng.Value = val;
            }
        }

        public static void SetCellFormula(Worksheet sheet, int row, int column, string formula)
        {
            formula = formula.Replace("|", "\"");
            formula = formula.Replace("~NE~", "<>");
            formula = formula.Replace("~GTE~", ">=");
            formula = formula.Replace("~LTE~", "<=");
            formula = formula.Replace("~LT~", "<");
            formula = formula.Replace("~GT~", ">");
            if (sheet != null)
            {
                Range rng = sheet.Cells[row, column] as Range;
                if (rng != null)
                    sheet.Cells[row, column].Formula = string.Format(formula, 1);
            }
        }

        public static string GetNewFileName(string summary, string epicID)
        {
            return @ThisAddIn.OutputDir + "\\" + "Epic ID " + epicID.Trim() + " " + GetValidFileName(summary.Trim() +  ".docx");
        }

        public static string GetValidFileName(string text)
        {
            text = text.Replace('\'', ' '); // U+2019 right single quotation mark
            text = text.Replace('"', ' '); // U+201D right double quotation mark
            text = text.Replace('/', ' ');  // U+2044 fraction slash
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                text = text.Replace(c, ' ');
            }
            return text;
        }

        public static void SetStandardRowHeight(Worksheet ws, int headerRow, int footerRow)
        {
            Excel.Range allRows = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
            allRows.EntireRow.RowHeight = 15;
        }

        //// http://www.authorcode.com/search-text-in-excel-file-through-c/
        public static int FindTextInColumn(Worksheet ws, string colRangeName, string valueToFind)
        {
            try
            {
                Excel.Range r = GetSpecifiedRange(valueToFind, ws, colRangeName);
                if (r != null)
                {
                    return r.Row;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return 0;
            }
        }

        public static Range GetSpecifiedRange(string valueToFind, Worksheet ws, string namedRange)
        {
            Excel.Range currentFind = null;
            currentFind = ws.get_Range(namedRange).Find(valueToFind, Missing.Value,
                           Excel.XlFindLookIn.xlValues,
                           Excel.XlLookAt.xlPart,
                           Excel.XlSearchOrder.xlByRows,
                           Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);
            return currentFind;
        }

        public static void DoStandardStuff(Excel.Application app)
        {
            //if (app.ScreenUpdating == true)
            //{
            //    app.Cursor = XlMousePointer.xlWait;
            //    app.ScreenUpdating = false;
            //    app.Calculation = XlCalculation.xlCalculationManual;
            //}
            //else
            //{
            //    app.Calculation = XlCalculation.xlCalculationAutomatic;
            //    app.ScreenUpdating = true;
            //    app.Cursor = XlMousePointer.xlDefault;
            //}
        }
    }
}
