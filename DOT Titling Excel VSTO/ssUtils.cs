using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class SSUtils
    {
        public static int GetFooterRow(Excel.Worksheet ws)
        {
            string sFooterRowRange = GetFooterRangeName(ws.Name);
            Excel.Range footerRangeRange = ws.get_Range(sFooterRowRange, Type.Missing);
            int footerRow = footerRangeRange.Row;
            return footerRow;
        }

        public static int GetHeaderRow(Excel.Worksheet ws)
        {
            string sHeaderRangeName = GetHeaderRangeName(ws.Name);
            Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;
            return headerRow;
        }

        public static string ZeroIfEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? "0" : s;
        }

        public static bool CheckDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string MissingColumns(Excel.Worksheet ws)
        {
            string missingFields = string.Empty;
            var jiraFields = WorksheetPropertiesManager.GetJiraFields(ws);
            foreach (var jiraField in jiraFields)
            {
                string columnHeader = jiraField.ColumnHeader;
                if (GetColumnFromHeader(ws, columnHeader) == 0)
                    missingFields = missingFields + ' ' + columnHeader;
            }
            return missingFields.Trim();
        }

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
            Excel.Range colRange = app.get_Range(tableRangeName + "[" + columnHeader + "]", Type.Missing);
            if (colRange != null)
                return GetColumnName(colRange.Column);
            return string.Empty;
        }

        public static void HideTableColumns(Excel.Range headerRowRange, List<string> ColumnsToShow)
        {
            // Format each cell in the table header row
            foreach (Excel.Range cell in headerRowRange.Cells)
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
            Excel.Worksheet ws = app.ActiveSheet;
            foreach (Excel.ListObject table in ws.ListObjects)
            {
                Excel.Range tableRange = table.Range;
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


        public static Excel.ListObject GetListObjectFromTableName(Excel.Worksheet ws, string tableName)
        {
            Excel.ListObject lo = null;
            foreach (Excel.ListObject table in ws.ListObjects)
            {
                Excel.Range tableRange = table.Range;
                if (table.Active == true && table.Name == tableName)
                {
                    lo = table;
                }
            }
            return lo;
        }

        public static void SetColumnWidth(Excel.Worksheet ws, string columnHeader, int width)
        {
            try
            {
                int col = GetColumnFromHeader(ws, columnHeader);
                if (col != 0)
                {
                    Excel.Range cell = ws.Cells[1, col] as Excel.Range;
                    cell.EntireColumn.ColumnWidth = width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static string GetSelectedTableFooter(Excel.Application app)
        {
            string f = string.Empty;
            Excel.Worksheet ws = app.ActiveSheet;
            string tableName = GetSelectedTable(app);
            if (tableName != string.Empty)
                f = tableName + "[#Totals]";
            return f;
        }

        public static List<string> GetListOfTables(Excel.Application app)
        {
            List<string> listofTables = new List<string>();
            Excel.Worksheet ws = app.ActiveSheet;
            foreach (Excel.ListObject table in ws.ListObjects)
            {
                listofTables.Add(table.Name);
                Excel.Range tableRange = table.Range;
                if (table.Active == true)
                    MessageBox.Show(table.Name);
            }
            return listofTables;
        }

        public static int GetColumnFromHeader(Excel.Worksheet ws, string columnText)
        {
            try
            {
                string sHeaderRangeName = GetHeaderRangeName(ws.Name);
                Excel.Range headerRowRange = ws.get_Range(sHeaderRangeName, Type.Missing);
                foreach (Excel.Range cell in headerRowRange.Cells)
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

        public static string GetCellValue(Excel.Worksheet sheet, int row, int column)
        {
            var result = string.Empty;
            if (sheet != null)
            {
                Excel.Range rng = sheet.Cells[row, column] as Excel.Range;

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
                Excel.Range rng = app.get_Range(rangeName);
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

        public static void SetCellValue(Excel.Worksheet sheet, int row, int column, string val)
        {
            try
            {
                if (sheet != null)
                {
                    if (column != 0)
                    {
                        Excel.Range rng = sheet.Cells[row, column] as Excel.Range;
                        if (rng != null)
                            rng.Value = val;
                    }
                    else
                    {
                        //MessageBox.Show(columnHeader + " is missing");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void SetCellFormula(Excel.Worksheet sheet, int row, int column, string formula)
        {
            formula = formula.Replace("|", "\"");
            formula = formula.Replace("~NE~", "<>");
            formula = formula.Replace("~GTE~", ">=");
            formula = formula.Replace("~LTE~", "<=");
            formula = formula.Replace("~LT~", "<");
            formula = formula.Replace("~GT~", ">");
            if (sheet != null)
            {
                Excel.Range rng = sheet.Cells[row, column] as Excel.Range;
                if (rng != null)
                    sheet.Cells[row, column].Formula = string.Format(formula, 1);
            }
        }

        public static void SetStandardRowHeight(Excel.Worksheet ws, int headerRow, int footerRow)
        {
            Excel.Range allRows = null;
            if (headerRow == footerRow)
            {
                allRows = ws.get_Range(String.Format("{0}:{1}", headerRow, footerRow), Type.Missing);
            }
            else
            { 
                allRows = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
            }
            allRows.EntireRow.RowHeight = 15;
        }

        //// http://www.authorcode.com/search-text-in-excel-file-through-c/
        public static int FindTextInColumn(Excel.Worksheet ws, string colRangeName, string valueToFind)
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

        public static Excel.Range GetSpecifiedRange(string valueToFind, Excel.Worksheet ws, string namedRange)
        {
            Excel.Range currentFind = null;
            currentFind = ws.get_Range(namedRange).Find(valueToFind, Missing.Value,
                           Excel.XlFindLookIn.xlValues,
                           Excel.XlLookAt.xlPart,
                           Excel.XlSearchOrder.xlByRows,
                           Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);
            return currentFind;
        }

        public async static Task<bool> BeginExcelOperation(Excel.Application app)
        {
            app.Cursor = Excel.XlMousePointer.xlWait;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;
            app.ScreenUpdating = false;
            return true;
        }

        public async static Task<bool> EndExcelOperation(Excel.Application app, string operationName)
        {
            app.Cursor = Excel.XlMousePointer.xlDefault;
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            app.ScreenUpdating = true;
            if (operationName != string.Empty)
                MessageBox.Show(operationName + " - Operation Complete");
            return true;
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
            //Excel.Worksheet ws = app.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "DOT Releases");
            //if (ws != null)
            //    app.Worksheets["DOT Releases"].Unprotect(Password: "dot333");
        }

        public static int GetLastRow(Excel.Worksheet ws)
        {
            return ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        public static int GetLastColumn(Excel.Worksheet ws)
        {
            return ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
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

        public static void FilterTable(Excel.Worksheet ws, string tableRangeName, string filterColumn, string filterValue)
        {
            try
            {
                Excel.ListObject list = GetListObjectFromTableName(ws, tableRangeName);
                list.AutoFilter.ShowAllData();
                int col = GetColumnFromHeader(ws, filterColumn);
                list.Range.AutoFilter(col, filterValue, Excel.XlAutoFilterOperator.xlFilterValues);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static Int32 TableRowCount(Excel.Worksheet ws, string tableRangeName)
        {
            try
            {
                Excel.ListObject list = GetListObjectFromTableName(ws, tableRangeName);
                return list.ListRows.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return 0;
            }
        }

        public static Int32 TableSelectedRowCount(Excel.Worksheet ws, Excel.Range selection)
        {
            try
            {
                Int32 rowCount = 0;
                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (ws.Rows[row].EntireRow.Height != 0)
                    {
                        rowCount++;
                    }
                }
                return rowCount;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return 0;
            }
        }

        public static void SortTable(Excel.Worksheet ws, string tableRangeName, string sortColumn, Excel.XlSortOrder sortOrder)
        {
            try
            {
                Excel.ListObject list = GetListObjectFromTableName(ws, tableRangeName);
                list.Range.Sort(
                    list.ListColumns[sortColumn].Range,
                    sortOrder,
                    list.ListColumns[2].Range,
                    Type.Missing,
                    Excel.XlSortOrder.xlAscending,
                    Type.Missing,
                    Excel.XlSortOrder.xlAscending,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing,
                    Type.Missing,
                    Excel.XlSortOrientation.xlSortColumns,
                    Excel.XlSortMethod.xlPinYin,
                    Excel.XlSortDataOption.xlSortNormal,
                    Excel.XlSortDataOption.xlSortNormal,
                    Excel.XlSortDataOption.xlSortNormal);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public async static Task<List<string>> GetListOfProjects(Excel.Application app)
        {
            try
            {
                List<string> listofProjects = new List<string>();
                Excel.Worksheet wsProjects = app.Worksheets["Projects"];
                string sHeaderRangeName = GetHeaderRangeName(wsProjects.Name);
                Excel.Range headerRowRange = wsProjects.get_Range(sHeaderRangeName, Type.Missing);

                string sFooterRangeName = GetFooterRangeName(wsProjects.Name);
                Excel.Range footerRowRange = wsProjects.get_Range(sFooterRangeName, Type.Missing);

                int headerRow = headerRowRange.Row;
                int footerRow = footerRowRange.Row;

                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    int includeCol = GetColumnFromHeader(wsProjects, "Include");
                    string include = GetCellValue(wsProjects, row, includeCol).Trim();
                    if (include == "x")
                    {
                        int projectKeyCol = GetColumnFromHeader(wsProjects, "Project Key");
                        string projectKey = GetCellValue(wsProjects, row, projectKeyCol).Trim();
                        listofProjects.Add(projectKey);
                    }
                }
                return listofProjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in GetListOfProjects:" + ex);
                return null;
            }
        }
    }
}
