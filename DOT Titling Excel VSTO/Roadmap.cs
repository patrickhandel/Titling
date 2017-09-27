using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{ 
    class RoadMap
    {
        public static void ExecuteUpdateRoadMap()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                UpdateRoadMap(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateRoadMap(Excel.Application app)
        {
            try
            {
                Worksheet wsEpics = app.Sheets["Epics"];
                Worksheet wsRoadmap = app.Sheets["Road Map"];

                Int32 firstRow = 5;
                Int32 rmRow = firstRow;

                // Delete the rows row in the Road Map
                Int32 lastRow = SSUtils.GetLastRow(wsRoadmap);
                Int32 lastColumn = SSUtils.GetLastColumn(wsRoadmap);
                Range rToDelete = wsRoadmap.get_Range(String.Format("{0}:{1}", rmRow + 2, lastRow), Type.Missing);
                rToDelete.Delete();

                string releaseName = string.Empty;
                string epicName = string.Empty;
                Int32 sprintFrom = 0;
                Int32 sprintTo = 0;

                Int32 prevRelease = 0;
                Int32 lastSprint = 0;

                Int32 row = 1;

                List<Epic> epics = GetListOfEpics(wsEpics);
                Int32 epicCount = epics.Count;

                foreach (var epic in epics)
                {
                    Int32 release = epic.Release;
                    if (release != 99)
                    {
                        releaseName = epic.ReleaseName;
                        epicName = epic.EpicName;
                        sprintFrom = epic.SprintFrom;
                        sprintTo = epic.SprintTo;

                        if (release != prevRelease)
                        {
                            if (row != 1)
                            {
                                CreateRow(wsRoadmap, "UAT", rmRow, prevRelease, lastSprint, releaseName, lastColumn, "", 0, 0);
                                rmRow++;
                            }

                            CreateRow(wsRoadmap, "REL", rmRow, prevRelease, lastSprint, releaseName, lastColumn, "", 0, 0);
                            rmRow++;

                            if (row != 1)
                            {
                                CreateRow(wsRoadmap, "BFP", rmRow, prevRelease, lastSprint, releaseName, lastColumn, "", 0, 0);
                                rmRow++;
                            }
                        }

                        CreateRow(wsRoadmap, "EPIC", rmRow, prevRelease, lastSprint, releaseName, lastColumn, epicName, sprintFrom, sprintTo);
                        rmRow++;
                        row++;
                        lastSprint = sprintTo;
                        prevRelease = release;
                    }
                }

                //releaseName = SSUtils.GetCellValue(wsEpics, row, releaseNameColumn);
                //sprintFrom = SSUtils.GetCellValue(wsEpics, row, sprintFromColumn);
                //sprintTo = SSUtils.GetCellValue(wsEpics, row, sprintToColumn);
                //epicName = SSUtils.GetCellValue(wsEpics, row, epicColumn);

                //CreateRow(wsRoadmap, "UAT", rmRow, prevRelease, lastSprint, releaseName, lastColumn, epicName, sprintFrom, sprintTo);
                //rmRow++;

                //CreateRow(wsRoadmap, "REL", rmRow, prevRelease, lastSprint, "Final Release", lastColumn, epicName, sprintFrom, sprintTo);
                //rmRow++;

                //CreateRow(wsRoadmap, "BFP", rmRow, prevRelease, lastSprint, releaseName, lastColumn, epicName, sprintFrom, sprintTo);
                //rmRow++;

                //CreateRow(wsRoadmap, "FINAL ROW", rmRow, prevRelease, lastSprint, releaseName, lastColumn, epicName, sprintFrom, sprintTo);
                //rmRow++;

                FormatChart(wsRoadmap, firstRow, rmRow, lastColumn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void FormatChart(Worksheet wsRoadmap, int firstRow, int rmRow, int lastColumn)
        {
            string range1val = string.Format("A{0}:{1}5", firstRow, ColumnNumberToName(lastColumn));
            string range2val = string.Format("A{0}:{1}{2}", firstRow + 1, ColumnNumberToName(lastColumn), rmRow - 1);
            string rangeToSelect = string.Format("A{0}", firstRow);
            Range range1 = wsRoadmap.get_Range(range1val);
            Range range2 = wsRoadmap.get_Range(range2val);
            Range range3 = wsRoadmap.get_Range(rangeToSelect);
            range1.Copy(Type.Missing);
            range2.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            range3.Select();
        }

        private static string ColumnNumberToName(Int32 col_num)
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

        private static Int32 ColumnNameToNumber(string col_name)
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

        private static List<Epic> GetListOfEpics(Worksheet wsEpics)
        {
            string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsEpics.Name);
            Range headerRowRange = wsEpics.get_Range(sHeaderRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            string sFooterRowRange = SSUtils.GetFooterRangeName(wsEpics.Name);
            Range footerRangeRange = wsEpics.get_Range(sFooterRowRange, Type.Missing);
            int footerRow = footerRangeRange.Row;

            int priorityColumn = SSUtils.GetColumnFromHeader(wsEpics, "Priority");
            int releaseColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release");
            int releaseNameColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release Name");
            int sprintFromColumn = SSUtils.GetColumnFromHeader(wsEpics, "Sprint From");
            int sprintToColumn = SSUtils.GetColumnFromHeader(wsEpics, "Sprint From");
            int epicColumn = SSUtils.GetColumnFromHeader(wsEpics, "Epic");

            var epics = new List<Epic>();
            for (int row = headerRow + 1; row < footerRow; row++)
            {
                string priority = SSUtils.GetCellValue(wsEpics, row, priorityColumn);
                string release = SSUtils.GetCellValue(wsEpics, row, releaseColumn);
                string releaseName = SSUtils.GetCellValue(wsEpics, row, releaseNameColumn);
                string sprintFrom = SSUtils.GetCellValue(wsEpics, row, sprintFromColumn);
                string sprintTo = SSUtils.GetCellValue(wsEpics, row, sprintToColumn);
                string epicName = SSUtils.GetCellValue(wsEpics, row, epicColumn);
                epics.Add(new Epic(epicName, releaseName, Convert.ToInt32(release), Convert.ToInt32(sprintFrom), Convert.ToInt32(sprintTo), Convert.ToInt32(priority)));
            }
            epics.Sort();
            return epics;
        }

        private static void CreateRow(Worksheet ws, string rowType, Int32 row, Int32 prevRelease, Int32 lastSprint, string releaseName, Int32 lastColumn, string epicName, Int32 sprintFrom, Int32 sprintTo)
        {
            try
            {
                SSUtils.SetCellValue(ws, row, 2, rowType);
                switch (rowType)
                {
                    case "UAT":
                        SSUtils.SetCellValue(ws, row, 1, "R" + prevRelease + " Release and UAT");
                        SSUtils.SetCellValue(ws, row, 3, (lastSprint + 2).ToString());
                        SSUtils.SetCellValue(ws, row, 4, (lastSprint + 2).ToString());
                        break;
                    case "REL":
                        SSUtils.SetCellValue(ws, row, 1, releaseName);
                        break;
                    case "BFP":
                        SSUtils.SetCellValue(ws, row, 1, "Bug Fixing Period - R" + prevRelease);
                        SSUtils.SetCellValue(ws, row, 3, (lastSprint + 2).ToString());
                        SSUtils.SetCellValue(ws, row, 4, (lastSprint + 3).ToString());
                        break;
                    case "EPIC":
                        SSUtils.SetCellValue(ws, row, 1, epicName);
                        SSUtils.SetCellValue(ws, row, 3, (sprintFrom).ToString());
                        SSUtils.SetCellValue(ws, row, 4, (sprintTo).ToString());
                        break;
                    case "FINAL ROW":
                        SSUtils.SetCellValue(ws, row, 2, "UAT");
                        SSUtils.SetCellValue(ws, row, 1, "Final Release");
                        SSUtils.SetCellValue(ws, row, 2, "UAT");
                        SSUtils.SetCellValue(ws, row, 3, (lastSprint + 4).ToString());
                        SSUtils.SetCellValue(ws, row, 4, (lastSprint + 5).ToString());
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


    }
}
