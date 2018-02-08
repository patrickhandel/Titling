using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class Lists
    {
        public static List<Issue> GetListOfIssues(Excel.Worksheet wsIssues)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsIssues.Name);
                Excel.Range headerRowRange = wsIssues.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsIssues.Name);
                Excel.Range footerRangeRange = wsIssues.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int idColumn = SSUtils.GetColumnFromHeader(wsIssues, "Issue ID");
                int typeColumn = SSUtils.GetColumnFromHeader(wsIssues, "Issue Type");
                int summaryColumn = SSUtils.GetColumnFromHeader(wsIssues, "Summary (Local)");
                int statusColumn = SSUtils.GetColumnFromHeader(wsIssues, "Status");
                int sprintColumn = SSUtils.GetColumnFromHeader(wsIssues, "DOT Sprint Number (Local)");

                var issues = new List<Issue>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string id = SSUtils.GetCellValue(wsIssues, row, idColumn);
                    string type = SSUtils.GetCellValue(wsIssues, row, typeColumn);
                    string summary = SSUtils.GetCellValue(wsIssues, row, summaryColumn);
                    string status = SSUtils.GetCellValue(wsIssues, row, statusColumn);
                    string sprint = SSUtils.GetCellValue(wsIssues, row, sprintColumn);
                    issues.Add(new Issue(
                        id,
                        type,
                        summary,
                        status,
                        sprint
                        ));
                }
                issues.Sort();
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static List<Release> GetListOfReleases(Excel.Worksheet wsReleases)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsReleases.Name);
                Excel.Range headerRowRange = wsReleases.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsReleases.Name);
                Excel.Range footerRangeRange = wsReleases.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int numberColumn = SSUtils.GetColumnFromHeader(wsReleases, "Release (Local)");
                int nameColumn = SSUtils.GetColumnFromHeader(wsReleases, "Full Name");
                int statusColumn = SSUtils.GetColumnFromHeader(wsReleases, "Status");
                int midLongColumn = SSUtils.GetColumnFromHeader(wsReleases, "Mid/Long");
                int sprintFromColumn = SSUtils.GetColumnFromHeader(wsReleases, "From");
                int sprintToColumn = SSUtils.GetColumnFromHeader(wsReleases, "To");
                int uatSprintFromColumn = SSUtils.GetColumnFromHeader(wsReleases, "UAT From");
                int uatSprintToColumn = SSUtils.GetColumnFromHeader(wsReleases, "UAT To");
                int vendorSprintColumn = SSUtils.GetColumnFromHeader(wsReleases, "Vendor Sprint");

                var releases = new List<Release>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string number = SSUtils.GetCellValue(wsReleases, row, numberColumn);
                    string name = SSUtils.GetCellValue(wsReleases, row, nameColumn);
                    string midLong = SSUtils.GetCellValue(wsReleases, row, midLongColumn);
                    string sprintFrom = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, sprintFromColumn));
                    string sprintTo = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, sprintToColumn));
                    string uatSprintFrom = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintFromColumn));
                    string uatSprintTo = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintToColumn));
                    string vendorSprint = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, vendorSprintColumn));
                    string status = SSUtils.GetCellValue(wsReleases, row, statusColumn);
                    releases.Add(new Release(
                        Convert.ToInt32(number),
                        name,
                        midLong,
                        Convert.ToInt32(sprintFrom),
                        Convert.ToInt32(sprintTo),
                        Convert.ToInt32(uatSprintFrom),
                        Convert.ToInt32(uatSprintTo),
                        Convert.ToInt32(vendorSprint),
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

        public static List<Epic> GetListOfEpics(Excel.Worksheet wsEpics)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsEpics.Name);
                Excel.Range headerRowRange = wsEpics.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsEpics.Name);
                Excel.Range footerRangeRange = wsEpics.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int priorityColumn = SSUtils.GetColumnFromHeader(wsEpics, "Priority");
                int releaseNumberColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release (Local)");
                int releaseNameColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release Name");
                int epicColumn = SSUtils.GetColumnFromHeader(wsEpics, "Epic (Local)");
                int midLongColumn = SSUtils.GetColumnFromHeader(wsEpics, "Mid/Long");
                int statusColumn = SSUtils.GetColumnFromHeader(wsEpics, "Percent Complete");

                var epics = new List<Epic>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string priority = SSUtils.GetCellValue(wsEpics, row, priorityColumn);
                    string releaseNumber = SSUtils.ZeroIfEmpty(SSUtils.GetCellValue(wsEpics, row, releaseNumberColumn));
                    string releaseName = SSUtils.GetCellValue(wsEpics, row, releaseNameColumn);
                    string epicName = SSUtils.GetCellValue(wsEpics, row, epicColumn);
                    string midLong = SSUtils.GetCellValue(wsEpics, row, midLongColumn);
                    string status = SSUtils.GetCellValue(wsEpics, row, statusColumn);
                    epics.Add(new Epic(
                            epicName,
                            releaseName,
                            Convert.ToInt32(releaseNumber),
                            Convert.ToInt32(priority),
                            status
                        ));
                }
                epics.Sort();
                return epics;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }
    }
}
