﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    class SSUtils
    {
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
                Excel.Range rng = sheet.Cells[row, column] as Excel.Range;

                if (rng != null)
                    result = (string)rng.Text;
            }
            return (result + " ").Trim();
        }

        public static void SetCellValue(Excel.Worksheet sheet, int row, int column, string val)
        {
            if (sheet != null)
            {
                Excel.Range rng = sheet.Cells[row, column] as Excel.Range;
                if (rng != null)
                    rng.Value = val;
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

        //// http://www.authorcode.com/search-text-in-excel-file-through-c/
        public static void SearchText(Excel.Application app, Excel.Workbook wb, Excel.Worksheet ws)
        {
            try
            {
                Excel.Range r = GetSpecifiedRange("test", ws);
                if (r != null)
                {
                    MessageBox.Show("Text found, position is Row:" + r.Row + " and column:" + r.Column);
                }
                else
                {
                    MessageBox.Show("Text is not found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static Excel.Range GetSpecifiedRange(string matchStr, Excel.Worksheet ws)
        {
            Excel.Range currentFind = null;
            currentFind = ws.get_Range("A1", "AM100").Find(matchStr, Missing.Value,
                           Excel.XlFindLookIn.xlValues,
                           Excel.XlLookAt.xlPart,
                           Excel.XlSearchOrder.xlByRows,
                           Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);
            return currentFind;
        }

        public static void SetStandardRowHeight(Excel.Worksheet ws, int headerRow, int footerRow)
        {
            Excel.Range allRows = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
            allRows.EntireRow.RowHeight = 15;
        }
    }
}
