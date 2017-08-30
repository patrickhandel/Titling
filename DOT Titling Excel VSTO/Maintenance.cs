using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace DOT_Titling_Excel_VSTO
{
    class Maintenance
    {
        public static void AddNewStories()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;

                if (activeWorksheet.Name == "Stories" || activeWorksheet.Name == "Jira Stories")
                {
                    //app.ScreenUpdating = false;
                    Excel.Worksheet wsStories = app.ActiveWorkbook.Sheets["Stories"];
                    Excel.Worksheet wsJiraStories = app.ActiveWorkbook.Sheets["Jira Stories"];
                    string sColRangeName = "JiraStoryData[ERR Found (WIN)]";
                    Excel.Range errColRange = (Excel.Range)wsJiraStories.get_Range(sColRangeName, Type.Missing);
                    string val = "";
                    foreach (Excel.Range cell in errColRange.Cells)
                    {
                        val = cell.Value;
                        if (val == "x")
                        {
                            int row = cell.Row;
                            string epic = SSUtils.GetCellValue(wsJiraStories, row, 22);
                            string id = SSUtils.GetCellValue(wsJiraStories, row, 1);
                            string summary = SSUtils.GetCellValue(wsJiraStories, row,3);

                            string sStoriesRange = "StoryData";
                            Excel.ListObject list = wsStories.ListObjects[sStoriesRange];

                            Excel.Range tbl = wsStories.Range[sStoriesRange];

                            var lastUsedRow = tbl.get_End(XlDirection.xlDown).Row - 1;

                            CopyRowsDown(lastUsedRow, 5, tbl, wsStories);

                            //Excel.ListRow myRow = tbl.ListObject.ListRows.AddEx(Type.Missing, true);



                            //myRow.Range.Cells[1, 1] = epic;
                            //myRow.Range.Cells[1, 5] = summary;
                            //myRow.Range.Cells[1, 6] = id;

                            //list.Refresh();
                            

                            //foreach (Excel.Range r in myRow.Range.Cells)
                            //{
                            //    if (r.Column == 1)
                            //        r.Value = epic;

                            //    if (r.Column == 5)
                            //        r.Value = summary;

                            //    if (r.Column == 6)
                            //        r.Value = id;
                            //}
                        }
                    }
                    //app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


        public static int GetLastRow(Worksheet worksheet)
        {
            int lastUsedRow = 1;
            Range range = worksheet.UsedRange;
            for (int i = 1; i < range.Columns.Count; i++)
            {
                int lastRow = range.Rows.Count;
                for (int j = range.Rows.Count; j > 0; j--)
                {
                    if (lastUsedRow < lastRow)
                    {
                        lastRow = j;
                        if (!String.IsNullOrWhiteSpace(Convert.ToString((worksheet.Cells[j, i] as Range).Value)))
                        {
                            if (lastUsedRow < lastRow)
                                lastUsedRow = lastRow;
                            if (lastUsedRow == range.Rows.Count)
                                return lastUsedRow - 1;
                            break;
                        }
                    }
                    else
                        break;
                }
            }
            return lastUsedRow;
        }


        public static void CopyRowsDown(int startrow, int count, Excel.Range oRange, Excel.Worksheet oSheet)
        {
            oRange = oSheet.get_Range(String.Format("{0}:{0}", startrow), System.Type.Missing);
            oRange.Select();
            oRange.Copy();
            //oApp.Selection.Copy();

            oRange = oSheet.get_Range(String.Format("{0}:{1}", startrow + 1, startrow + count - 1), System.Type.Missing);
            oRange.Select();
            oRange.Insert(-4121);
            //oApp.Selection.Insert(-4121);

        }



    }


}
