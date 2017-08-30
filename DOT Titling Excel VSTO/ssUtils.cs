using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class SSUtils
    {
        public static int GetMailMergeFieldColumn(Excel.Worksheet ws, string columnText)
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
                return "";
            return prop.Range + "[#Headers]";
        }

        public static int GetColumnWidth(string name)
        {
            List<ColumnTypes> wsColumnTypes = WorksheetPropertiesManager.GetColumnTypes();
            var prop = wsColumnTypes.FirstOrDefault(p => p.Name == name);
            if (prop == null)
                return 15;
            return (int)prop.Width;
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
            return result + " ";
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
    }
}
