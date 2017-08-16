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
    class ssUtils
    {
        public static string GetCellValue(Excel.Worksheet sheet, int row, int column)
        {
            var result = string.Empty;
            if (sheet != null)
            {
                var rng = sheet.Cells[row, column] as Excel.Range;

                if (rng != null)
                    result = (string)rng.Text;
            }
            return result + " ";
        }

        public static string GetNewFileName(string summary, string id)
        {
            return @ThisAddIn.OutputDir + "\\" + GetValidFileName(summary.Trim() + " (" + id.Trim() + ").docx");
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
