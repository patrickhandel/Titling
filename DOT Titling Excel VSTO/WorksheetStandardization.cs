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
                        string colType = cell.Offset[-1, 0].Value;
                        switch (colType)
                        {
                            case "TextMedium":
                                i = ColumnType.TextMedium;
                                break;
                            case "TextLong":
                                i = ColumnType.TextLong;
                                break;
                            case "TextShort":
                                i = ColumnType.TextShort;
                                break;
                            case "Number":
                                i = ColumnType.Number;
                                break;
                            case "YesNo":
                                i = ColumnType.YesNo;
                                break;
                            case "Percent":
                                i = ColumnType.Percent;
                                break;
                            case "Error":
                                i = ColumnType.Error;
                                break;
                            case "Date":
                                i = ColumnType.Date;
                                break;
                            default:
                                i = ColumnType.Default;
                                break;
                        }
                        cell.EntireColumn.ColumnWidth = i;
                    }

                    Excel.Range r = ws.get_Range("A1");
                    r.EntireRow.RowHeight = 60;

                    headerRowRange.EntireRow.RowHeight = 60;
 
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
            string rn;    
            if (name == "Stories")
            {
                rn = "StoryData[#Headers]";
            }
            else if (name == "Jira Stories")
            {
                rn = "JiraStoryData[#Headers]";
            }
            else if (name == "Epics")
            {
                rn = "EpicData[#Headers]";
            }
            else if (name == "Jira Epics")
            {
                rn = "JiraEpicData[#Headers]";
            }
            else if (name == "Sprints")
            {
                rn = "SprintData[#Headers]";
            }
            else if (name == "Releases")
            {
                rn = "ReleaseData[#Headers]";
            }
            else if (name == "Jira Web Services")
            {
                rn = "JiraWebServicesData[#Headers]";
            }
            else if (name == "Sprint Results")
            {
                rn = "SprintResultsData[#Headers]";
            }
            else if (name == "Dev Results")
            {
                rn = "DevResultsData[#Headers]";
            }
            else if (name == "Jira Bugs")
            {
                rn = "JiraBugData[#Headers]";
            }
            else
            {
                rn = "";
            }

            return rn;
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
        public const int Error = 8;
        public const int Date = 11;
        public const int Default = 15;
    }
}
