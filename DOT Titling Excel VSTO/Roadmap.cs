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
                Worksheet wsReleases = app.Sheets["Releases"];
                Worksheet wsRoadmap = app.Sheets["Road Map"];

                Int32 firstRow = 5;
                Int32 rmRow = firstRow;

                // Delete the rows row in the Road Map
                Int32 lastRow = SSUtils.GetLastRow(wsRoadmap);
                Int32 lastColumn = SSUtils.GetLastColumn(wsRoadmap);

                if (lastRow > rmRow + 2)
                {
                    Range rToDelete = wsRoadmap.get_Range(String.Format("{0}:{1}", rmRow + 2, lastRow), Type.Missing);
                    rToDelete.Delete();
                }

                List<Release> releases = GetListOfReleases(wsReleases);
                List<Epic> epics = GetListOfEpics(wsEpics);

                string prevReleaseName = string.Empty;
                Int32 prevSprintTo = 0;
                bool firstRelease = true;
                foreach (var release in releases)
                {
                    Int32 releaseNumber = release.Number;
                    string releaseName = release.Name;
                    string releaseStatus = release.Status;
                    Int32 sprintFrom = release.DevSprintFrom;
                    Int32 sprintTo = release.DevSprintTo;
                    Int32 uatSprintFrom = release.UATSprintFrom;
                    Int32 uatSprintTo = release.UATSprintTo;

                    List<Epic> releaseEpics = epics.FindAll(e => e.ReleaseName == release.Name && e.MidLong == "Mid");
                    if (releaseEpics.Count > 0)
                    {
                        // REL
                        CreateRow(wsRoadmap, "REL", rmRow, "", releaseName, releaseStatus, releaseNumber, 0, 0);
                        rmRow++;

                        // BFP
                        if (!firstRelease)
                        {
                            CreateRow(wsRoadmap, "BFP", rmRow, "", releaseName, "", releaseNumber - 1, prevSprintTo + 1, prevSprintTo + 2);
                            rmRow++;
                        }

                        foreach (var epic in releaseEpics)
                        {
                            // EPIC
                            string epicName = epic.EpicName;
                            CreateRow(wsRoadmap, "EPIC", rmRow, epicName, releaseName, epic.Status, releaseNumber,  sprintFrom, sprintTo);
                            rmRow++;
                        }

                        // UAT
                        CreateRow(wsRoadmap, "UAT", rmRow, "", releaseName, "", releaseNumber, uatSprintFrom, uatSprintTo);
                        rmRow++;
                    }
                    prevReleaseName = releaseName;
                    prevSprintTo = sprintTo;
                    firstRelease = false;
                }
                FormatChart(wsRoadmap, firstRow, rmRow, lastColumn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void CreateRow(Worksheet ws, string rowType, Int32 row, string epicName, string releaseName, string status, Int32 releaseNumber, Int32 sprintFrom, Int32 sprintTo)
        {
            try
            {
                switch (rowType)
                {
                    case "BFP":
                        SSUtils.SetCellValue(ws, row, 1, "Bug Fixing - R" + releaseNumber.ToString());
                        SSUtils.SetCellValue(ws, row, 2, string.Empty);
                        SSUtils.SetCellValue(ws, row, 3, rowType);
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString());
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString());
                        break;
                    case "UAT":
                        SSUtils.SetCellValue(ws, row, 1, "R" + releaseNumber.ToString() + " Release and UAT");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty);
                        SSUtils.SetCellValue(ws, row, 3, rowType);
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString());
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString());
                        break;
                    case "REL":
                        SSUtils.SetCellValue(ws, row, 1, releaseName);
                        SSUtils.SetCellValue(ws, row, 2, status);
                        SSUtils.SetCellValue(ws, row, 3, rowType);
                        SSUtils.SetCellValue(ws, row, 4, string.Empty);
                        SSUtils.SetCellValue(ws, row, 5, string.Empty);
                        break;
                    case "EPIC":
                        SSUtils.SetCellValue(ws, row, 1, epicName);
                        SSUtils.SetCellValue(ws, row, 2, status);
                        SSUtils.SetCellValue(ws, row, 3, rowType);
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString());
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString());
                        break;
                    case "FINAL ROW":
                        SSUtils.SetCellValue(ws, row, 1, "FINAL ROW");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty);
                        SSUtils.SetCellValue(ws, row, 3, string.Empty);
                        SSUtils.SetCellValue(ws, row, 4, string.Empty);
                        SSUtils.SetCellValue(ws, row, 5, string.Empty);
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

        private static void FormatChart(Worksheet wsRoadmap, int firstRow, int rmRow, int lastColumn)
        {
            string range1val = string.Format("A{0}:{1}5", firstRow, SSUtils.ColumnNumberToName(lastColumn));
            string range2val = string.Format("A{0}:{1}{2}", firstRow + 1, SSUtils.ColumnNumberToName(lastColumn), rmRow - 1);
            string rangeToSelect = string.Format("A{0}", firstRow);
            Range range1 = wsRoadmap.get_Range(range1val);
            Range range2 = wsRoadmap.get_Range(range2val);
            Range range3 = wsRoadmap.get_Range(rangeToSelect);
            range1.Copy(Type.Missing);
            range2.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            range3.Select();
        }

        private static List<Release> GetListOfReleases(Worksheet wsReleases)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsReleases.Name);
                Range headerRowRange = wsReleases.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsReleases.Name);
                Range footerRangeRange = wsReleases.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int numberColumn = SSUtils.GetColumnFromHeader(wsReleases, "Release");
                int nameColumn = SSUtils.GetColumnFromHeader(wsReleases, "Full Name");
                int devSprintFromColumn = SSUtils.GetColumnFromHeader(wsReleases, "Dev Sprint (From)");
                int devSprintToColumn = SSUtils.GetColumnFromHeader(wsReleases, "Dev Sprint (To)");
                int uatSprintFromColumn = SSUtils.GetColumnFromHeader(wsReleases, "UAT Sprint (From)");
                int uatSprintToColumn = SSUtils.GetColumnFromHeader(wsReleases, "UAT Sprint (To)");
                int statusColumn = SSUtils.GetColumnFromHeader(wsReleases, "Status");

                var releases = new List<Release>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string number = SSUtils.GetCellValue(wsReleases, row, numberColumn);
                    string name = SSUtils.GetCellValue(wsReleases, row, nameColumn);
                    string devSprintFrom = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, devSprintFromColumn));
                    string devSprintTo = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, devSprintToColumn));
                    string uatSprintFrom = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintFromColumn));
                    string uatSprintTo = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintToColumn));
                    string status = SSUtils.GetCellValue(wsReleases, row, statusColumn);
                    releases.Add(new Release(
                        Convert.ToInt32(number), 
                        name, 
                        Convert.ToInt32(devSprintFrom), 
                        Convert.ToInt32(devSprintTo), 
                        Convert.ToInt32(uatSprintFrom), 
                        Convert.ToInt32(uatSprintTo),
                        status));
                }
                releases.Sort();
                return releases;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
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
            int midLongColumn = SSUtils.GetColumnFromHeader(wsEpics, "Mid/Long");
            int statusColumn = SSUtils.GetColumnFromHeader(wsEpics, "Percent Complete");

            var epics = new List<Epic>();
            for (int row = headerRow + 1; row < footerRow; row++)
            {
                string priority = SSUtils.GetCellValue(wsEpics, row, priorityColumn);
                string release = ZeroIfEmpty(SSUtils.GetCellValue(wsEpics, row, releaseColumn));
                string releaseName = SSUtils.GetCellValue(wsEpics, row, releaseNameColumn);
                string sprintFrom = ZeroIfEmpty(SSUtils.GetCellValue(wsEpics, row, sprintFromColumn));
                string sprintTo = ZeroIfEmpty(SSUtils.GetCellValue(wsEpics, row, sprintToColumn));
                string epicName = SSUtils.GetCellValue(wsEpics, row, epicColumn);
                string midLong = SSUtils.GetCellValue(wsEpics, row, midLongColumn);
                string status = SSUtils.GetCellValue(wsEpics, row, statusColumn);
                epics.Add(new Epic(epicName, releaseName, Convert.ToInt32(release), Convert.ToInt32(sprintFrom), Convert.ToInt32(sprintTo), Convert.ToInt32(priority), midLong, status));
            }
            epics.Sort();
            return epics;
        }

        private static string ZeroIfEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? "0" : s;
        }
    }
}
