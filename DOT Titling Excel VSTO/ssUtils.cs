using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class SSUtils
    {
        public static string GetColumnName(int columnNumber)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnName = "";

            while (columnNumber > 0)
            {
                columnName = letters[(columnNumber - 1) % 26] + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }
            return columnName;
        }

        public static string GetColumnLetter(Excel.Application app, string tableRangeName, string columnHeader)
        {
            Range colRange = app.get_Range(tableRangeName + "[" + columnHeader + "]", Type.Missing);
            if (colRange != null)
                return GetColumnName(colRange.Column);
            return string.Empty;
        }

        public static void HideTableColumns(Range headerRowRange, List<string> ColumnsToShow)
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

        public static string GetSelectedTable(Excel.Application app)
        {
            string t = string.Empty;
            Worksheet activeWorksheet = app.ActiveSheet;
            foreach (ListObject table in activeWorksheet.ListObjects)
            {
                Range tableRange = table.Range;
                if (table.Active == true)
                {
                    t = table.Name;
                    if (table.ShowTotals == false)
                        table.ShowTotals = true;
                }
            }
            return t;
        }

        public static string GetSelectedTableHeader(Excel.Application app)
        {
            string h = string.Empty;
            string tableName = GetSelectedTable(app);
            if (tableName != string.Empty)
                h = tableName + "[#Headers]";
            return h;
        }


        public static ListObject GetListObjectFromTableName(Worksheet ws, string tableName)
        {
            ListObject lo = null;
            foreach (ListObject table in ws.ListObjects)
            {
                Range tableRange = table.Range;
                if (table.Active == true && table.Name == tableName)
                {
                    lo = table;
                }
            }
            return lo;
        }

        public static string GetSelectedTableFooter(Excel.Application app)
        {
            string f = string.Empty;
            Excel.Worksheet activeWorksheet = app.ActiveSheet;
            string tableName = GetSelectedTable(app);
            if (tableName != string.Empty)
                f = tableName + "[#Totals]";
            return f;
        }

        public static List<string> GetListOfTables(Excel.Application app)
        {
            List<string> listofTables = new List<string>();
            Worksheet activeWorksheet = app.ActiveSheet;
            foreach (ListObject table in activeWorksheet.ListObjects)
            {
                listofTables.Add(table.Name);
                Range tableRange = table.Range;
                if (table.Active == true)
                    MessageBox.Show(table.Name);
            }
            return listofTables;
        }

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

        public static string GetCellValueFromNamedRange(string rangeName)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                var result = string.Empty;
                Range rng = app.get_Range(rangeName);
                if (rng != null)
                    result = (string)rng.Text;
                return result.Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return string.Empty;
            }
        }

        public static void SetCellValue(Worksheet sheet, int row, int column, string val, string columnHeader)
        {
            try
            {
                if (sheet != null)
                {
                    if (column != 0)
                    {

                        Range rng = sheet.Cells[row, column] as Range;
                        if (rng != null)
                            rng.Value = val;
                    }
                    else
                    {
                        MessageBox.Show(columnHeader + " is missing");
                    }
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
                Range r = GetSpecifiedRange(valueToFind, ws, colRangeName);
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

        public static void FilterTable(Worksheet ws, string tableRangeName, string filterColumn, string filterValue)
        {
            try
            {
                ListObject list = GetListObjectFromTableName(ws, tableRangeName);
                list.AutoFilter.ShowAllData();
                int col = GetColumnFromHeader(ws, filterColumn);
                list.Range.AutoFilter(col, filterValue, XlAutoFilterOperator.xlFilterValues);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void SortTable(Worksheet ws, string tableRangeName, string sortColumn, XlSortOrder sortOrder)
        {
            try
            {
                ListObject list = GetListObjectFromTableName(ws, tableRangeName);
                list.Range.Sort(
                    list.ListColumns[sortColumn].Range,
                    sortOrder,
                    list.ListColumns[2].Range,
                    Type.Missing, 
                    XlSortOrder.xlAscending,
                    Type.Missing, 
                    XlSortOrder.xlAscending,
                    XlYesNoGuess.xlYes,
                    Type.Missing,
                    Type.Missing,
                    XlSortOrientation.xlSortColumns,
                    XlSortMethod.xlPinYin,
                    XlSortDataOption.xlSortNormal,
                    XlSortDataOption.xlSortNormal,
                    XlSortDataOption.xlSortNormal);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


    }
}
