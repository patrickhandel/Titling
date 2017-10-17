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

        public static string GetCellValue(Worksheet sheet, int row, int column)
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
            try
            {
                if (sheet != null)
                {
                    Range rng = sheet.Cells[row, column] as Range;
                    if (rng != null)
                        rng.Value = val;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
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
            Range allRows = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
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
            Range currentFind = null;
            currentFind = ws.get_Range(namedRange).Find(valueToFind, Missing.Value,
                           Excel.XlFindLookIn.xlValues,
                           Excel.XlLookAt.xlPart,
                           Excel.XlSearchOrder.xlByRows,
                           Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);
            return currentFind;
        }

        public static void BeginExcelOperation(Excel.Application app)
        {
            app.Cursor = XlMousePointer.xlWait;
            app.Calculation = XlCalculation.xlCalculationManual;
            app.ScreenUpdating = false;
            UnProtect(app);
        }

        public static void EndExcelOperation(Excel.Application app, string operationName)
        {
            app.Cursor = XlMousePointer.xlDefault;
            app.Calculation = XlCalculation.xlCalculationAutomatic;
            app.ScreenUpdating = true;
            if (operationName != string.Empty)
                MessageBox.Show(operationName + " - Operation Complete");
            Protect(app);
        }

        private static void Protect(Excel.Application app)
        {
            //Worksheet ws = app.Worksheets.OfType<Worksheet>().FirstOrDefault(w => w.Name == "DOT Releases");
            //if (ws != null)
            //    app.Worksheets["DOT Releases"].Protect(Password: "dot333", 
            //            UserInterfaceOnly: false, 
            //            AllowFormattingCells: false, 
            //            AllowInsertingHyperlinks: false,
            //            AllowDeletingColumns: false, 
            //            AllowDeletingRows: false,
            //            AllowFormattingRows: true,
            //            AllowInsertingColumns: true,
            //            AllowInsertingRows: true,
            //            AllowSorting: true, 
            //            AllowFiltering: true,
            //            AllowFormattingColumns: true,
            //            AllowUsingPivotTables: true);
        }
        private static void UnProtect(Excel.Application app)
        {
            // https://msdn.microsoft.com/library/microsoft.office.interop.excel._worksheet.protect(v=office.15).aspx
            Worksheet ws = app.Worksheets.OfType<Worksheet>().FirstOrDefault(w => w.Name == "DOT Releases");
            if (ws != null)
                app.Worksheets["DOT Releases"].Unprotect(Password: "dot333");
        }

        public static int GetLastRow(Worksheet ws)
        {
            return ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        public static int GetLastColumn(Worksheet ws)
        {
            return ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
        }

        public static string ColumnNumberToName(Int32 col_num)
        {
            // See if it's out of bounds.
            if (col_num < 1) return "A";

            // Calculate the letters.
            string result = "";
            while (col_num > 0)
            {
                // Get the least significant digit.
                col_num -= 1;
                int digit = col_num % 26;

                // Convert the digit into a letter.
                result = (char)((Int32)'A' + digit) + result;

                col_num = (Int32)(col_num / 26);
            }

            return result;
        }

        public static Int32 ColumnNameToNumber(string col_name)
        {
            int result = 0;

            // Process each letter.
            for (Int32 i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (Int32)letter - (Int32)'A' + 1;
            }
            return result;
        }

        public static void SortTable(Excel.Application app, Worksheet ws, string rangeName, string column)
        {
            try
            {
                Range rng = ws.get_Range(rangeName);
                ListObject list = ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, rng, Type.Missing, XlYesNoGuess.xlYes, Type.Missing);
                list.Range.Sort(list.ListColumns[column].Range, XlSortOrder.xlAscending);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
