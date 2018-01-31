using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{ 
    class RoadMap
    {
        public static void ExecuteUpdateRoadMap_DOT(Excel.Application app)
        {
            try
            {
                var activeWorksheet = app.ActiveSheet;
                if (activeWorksheet.Name == "Road Map")
                {
                    UpdateRoadMap(app, activeWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateRoadMap(Excel.Application app, Worksheet wsRoadmap)
        {
            try
            {
                // Save into a PDF and Image.
                FileIO.CreateRoadMapImage(wsRoadmap);
                FileIO.CreateRoadMapPDF(wsRoadmap);

                Worksheet wsRoadMapBlocks = app.Sheets["Road Map Blocks"];
                FileIO.CreateRoadMapImage(wsRoadMapBlocks);
                FileIO.CreateRoadMapPDF(wsRoadMapBlocks);

                wsRoadmap.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateRoadMap_Ver1(Excel.Application app, Worksheet wsRoadmap)
        {
            try
            {
                Worksheet wsEpics = app.Sheets["Epics"];
                Worksheet wsReleases = app.Sheets["Releases"];

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

                List<Release> releases = Lists.GetListOfReleases(wsReleases);
                releases = releases.FindAll(r => r.MidLong == "Mid" || r.MidLong == "Long");
                List<Epic> epics = Lists.GetListOfEpics(wsEpics);

                string prevReleaseName = string.Empty;
                Int32 prevSprintTo = 0;
                bool firstRelease = true;
                foreach (var release in releases)
                {
                    Int32 releaseNumber = release.Number;
                    string releaseName = release.Name;
                    string midLong = release.MidLong;
                    string releaseStatus = release.Status;
                    Int32 sprintFrom = release.SprintFrom;
                    Int32 sprintTo = release.SprintTo;
                    Int32 uatSprintFrom = release.UATSprintFrom;
                    Int32 uatSprintTo = release.UATSprintTo;
                    Int32 vendorSprint = release.VendorSprint;

                    List<Epic> releaseEpics = epics.FindAll(e => e.ReleaseName == release.Name);
                    if (releaseEpics.Count > 0)
                    {
                        // REL
                        CreateRow(wsRoadmap, "REL", rmRow, "", releaseName, releaseStatus, releaseNumber, 0, 0, 0);
                        rmRow++;

                        // BFP
                        if (!firstRelease)
                        {
                            CreateRow(wsRoadmap, "BFP", rmRow, "", releaseName, "", releaseNumber - 1, prevSprintTo + 1, prevSprintTo + 2, 0);
                            rmRow++;
                        }

                        foreach (var epic in releaseEpics)
                        {
                            // EPIC
                            string epicName = epic.EpicName;
                            CreateRow(wsRoadmap, "EPIC", rmRow, epicName, releaseName, epic.Status, releaseNumber, sprintFrom, sprintTo, 0);
                            rmRow++;
                        }

                        // UAT
                        if (uatSprintTo != 0)
                        {
                            CreateRow(wsRoadmap, "UAT", rmRow, "", releaseName, "", releaseNumber, uatSprintFrom, uatSprintTo, 0);
                            rmRow++;
                        }

                        // VENDOR
                        if (vendorSprint != 0)
                        {
                            CreateRow(wsRoadmap, "VENDOR", rmRow, "", releaseName, "", releaseNumber, 0, 0, vendorSprint);
                            rmRow++;
                        }
                    }
                    prevReleaseName = releaseName;
                    prevSprintTo = sprintTo;
                    firstRelease = false;
                }
                FormatChart(wsRoadmap, firstRow, rmRow, lastColumn);

                // Save into a PDF and Image.
                FileIO.CreateRoadMapImage(wsRoadmap);
                FileIO.CreateRoadMapPDF(wsRoadmap);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void UpdateRoadMap_Ver2(Excel.Application app, Worksheet wsRoadmap)
        {
            try
            {
                Worksheet wsEpics = app.Sheets["Epics"];
                Worksheet wsReleases = app.Sheets["Releases"];

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

                List<Release> releases = Lists.GetListOfReleases(wsReleases);
                releases = releases.FindAll(r => r.MidLong == "Mid" || r.MidLong == "Long");

                string prevReleaseName = string.Empty;
                Int32 prevSprintTo = 0;
                bool firstRelease = true;
                foreach (var release in releases)
                {
                    Int32 releaseNumber = release.Number;
                    string releaseName = release.Name;
                    string midLong = release.MidLong;
                    string releaseStatus = release.Status;
                    Int32 sprintFrom = release.SprintFrom;
                    Int32 sprintTo = release.SprintTo;
                    Int32 uatSprintFrom = release.UATSprintFrom;
                    Int32 uatSprintTo = release.UATSprintTo;
                    Int32 vendorSprint = release.VendorSprint;

                    // REL
                    CreateRow(wsRoadmap, "REL", rmRow, "", releaseName, releaseStatus, releaseNumber, 0, 0, 0);
                    rmRow++;

                    // BFP
                    if (!firstRelease)
                    {
                        CreateRow(wsRoadmap, "BFP", rmRow, "", releaseName, "", releaseNumber - 1, prevSprintTo + 1, prevSprintTo + 2, 0);
                        rmRow++;
                    }

                    // EPIC
                    CreateRow(wsRoadmap, "DEV", rmRow, "", releaseName, "", releaseNumber, sprintFrom, sprintTo, 0);
                    rmRow++;

                    // UAT
                    if (uatSprintTo != 0)
                    {
                        CreateRow(wsRoadmap, "UAT", rmRow, "", releaseName, "", releaseNumber, uatSprintFrom, uatSprintTo, 0);
                        rmRow++;
                    }

                    // VENDOR
                    if (vendorSprint != 0)
                    {
                        CreateRow(wsRoadmap, "VENDOR", rmRow, "", releaseName, "", releaseNumber, 0, 0, vendorSprint);
                        rmRow++;
                    }

                    prevReleaseName = releaseName;
                    prevSprintTo = sprintTo;
                    firstRelease = false;
                }
                FormatChart(wsRoadmap, firstRow, rmRow, lastColumn);

                // Save into a PDF and Image.
                FileIO.CreateRoadMapImage(wsRoadmap);
                FileIO.CreateRoadMapPDF(wsRoadmap);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void CreateRow(Worksheet ws, string rowType, Int32 row, string epicName, string releaseName, string status, Int32 releaseNumber, Int32 sprintFrom, Int32 sprintTo, Int32 vendorSprint)
        {
            try
            {
                switch (rowType)
                {
                    case "BFP":
                        SSUtils.SetCellValue(ws, row, 1, "R" + releaseNumber.ToString() + " Bug Fixing", "?");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 3, rowType, "?");
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString(), "?");
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString(), "?");
                        break;
                    case "UAT":
                        SSUtils.SetCellValue(ws, row, 1, "R" + releaseNumber.ToString() + " Release and UAT", "?");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 3, rowType, "?");
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString(), "?");
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString(), "?");
                        break;
                    case "VENDOR":
                        SSUtils.SetCellValue(ws, row, 1, "R" + releaseNumber.ToString() + " Vendor Release", "?");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 3, rowType, "?");
                        SSUtils.SetCellValue(ws, row, 4, vendorSprint.ToString(), "?");
                        SSUtils.SetCellValue(ws, row, 5, vendorSprint.ToString(), "?");
                        break;
                    case "REL":
                        SSUtils.SetCellValue(ws, row, 1, releaseName, "?");
                        SSUtils.SetCellValue(ws, row, 2, status, "?");
                        SSUtils.SetCellValue(ws, row, 3, rowType, "?");
                        SSUtils.SetCellValue(ws, row, 4, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 5, string.Empty, "?");
                        break;
                    case "DEV":
                        SSUtils.SetCellValue(ws, row, 1, "R" + releaseNumber.ToString() + " Development", "?");
                        SSUtils.SetCellValue(ws, row, 2, status, "?");
                        SSUtils.SetCellValue(ws, row, 3, "Epic", "?");
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString(), "?");
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString(), "?");
                        break;
                    case "EPIC":
                        SSUtils.SetCellValue(ws, row, 1, epicName, "?");
                        SSUtils.SetCellValue(ws, row, 2, status, "?");
                        SSUtils.SetCellValue(ws, row, 3, rowType, "?");
                        SSUtils.SetCellValue(ws, row, 4, sprintFrom.ToString(), "?");
                        SSUtils.SetCellValue(ws, row, 5, sprintTo.ToString(), "?");
                        break;

                    case "FINAL ROW":
                        SSUtils.SetCellValue(ws, row, 1, "FINAL ROW", "?");
                        SSUtils.SetCellValue(ws, row, 2, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 3, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 4, string.Empty, "?");
                        SSUtils.SetCellValue(ws, row, 5, string.Empty, "?");
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

        private static string ZeroIfEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? "0" : s;
        }
    }
}
