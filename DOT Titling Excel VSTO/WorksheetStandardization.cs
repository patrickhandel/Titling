using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace DOT_Titling_Excel_VSTO
{
    class WorksheetStandardization
    {
        public static void ExecuteCleanup()
        {
            try
            {
                // Make sure this is a DOT workbook
                // Load the epics workbook
                // go through each column and find the header
                // if here is a header settings, set the column width and set the format
                // If it's an input field, confirm that there are no calculations

                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Workbook wb = app.ActiveWorkbook;
                Excel.Worksheet ws = wb.Sheets["Epics"];

                app.ScreenUpdating = false;

                int row = 4;
                int col = 1;
                bool finished = false;
                while (finished == false)
                {
                    string sVal = ssUtils.GetCellValue(ws, row, col).Trim();
                }

                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
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
    }
}
