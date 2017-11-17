using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows;

namespace DOT_Titling_Excel_VSTO
{
    class Lists
    {
        public static List<Ticket> GetListOfTickets(Worksheet wsTickets)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsTickets.Name);
                Range headerRowRange = wsTickets.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsTickets.Name);
                Range footerRangeRange = wsTickets.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int idColumn = SSUtils.GetColumnFromHeader(wsTickets, "Ticket ID");
                int typeColumn = SSUtils.GetColumnFromHeader(wsTickets, "Ticket Type");
                int summaryColumn = SSUtils.GetColumnFromHeader(wsTickets, "Summary");
                int statusColumn = SSUtils.GetColumnFromHeader(wsTickets, "Jira Status");
                int sprintColumn = SSUtils.GetColumnFromHeader(wsTickets, "DOT Sprint");

                var tickets = new List<Ticket>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string id = SSUtils.GetCellValue(wsTickets, row, idColumn);
                    string type = SSUtils.GetCellValue(wsTickets, row, typeColumn);
                    string summary = SSUtils.GetCellValue(wsTickets, row, summaryColumn);
                    string status = SSUtils.GetCellValue(wsTickets, row, statusColumn);
                    string sprint = SSUtils.GetCellValue(wsTickets, row, sprintColumn);
                    tickets.Add(new Ticket(
                        id,
                        type,
                        summary,
                        status,
                        sprint
                        ));

                }
                tickets.Sort();
                return tickets;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static List<Release> GetListOfReleases(Worksheet wsReleases)
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
                    string sprintFrom = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, sprintFromColumn));
                    string sprintTo = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, sprintToColumn));
                    string uatSprintFrom = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintFromColumn));
                    string uatSprintTo = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, uatSprintToColumn));
                    string vendorSprint = ZeroIfEmpty(SSUtils.GetCellValue(wsReleases, row, vendorSprintColumn));
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

        public static List<Epic> GetListOfEpics(Worksheet wsEpics)
        {
            try
            {
                string sHeaderRangeName = SSUtils.GetHeaderRangeName(wsEpics.Name);
                Range headerRowRange = wsEpics.get_Range(sHeaderRangeName, Type.Missing);
                int headerRow = headerRowRange.Row;

                string sFooterRowRange = SSUtils.GetFooterRangeName(wsEpics.Name);
                Range footerRangeRange = wsEpics.get_Range(sFooterRowRange, Type.Missing);
                int footerRow = footerRangeRange.Row;

                int priorityColumn = SSUtils.GetColumnFromHeader(wsEpics, "Priority");
                int releaseNumberColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release");
                int releaseNameColumn = SSUtils.GetColumnFromHeader(wsEpics, "Release Name");
                int epicColumn = SSUtils.GetColumnFromHeader(wsEpics, "Epic");
                int midLongColumn = SSUtils.GetColumnFromHeader(wsEpics, "Mid/Long");
                int statusColumn = SSUtils.GetColumnFromHeader(wsEpics, "Percent Complete");

                var epics = new List<Epic>();
                for (int row = headerRow + 1; row < footerRow; row++)
                {
                    string priority = SSUtils.GetCellValue(wsEpics, row, priorityColumn);
                    string releaseNumber = ZeroIfEmpty(SSUtils.GetCellValue(wsEpics, row, releaseNumberColumn));
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
        private static string ZeroIfEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? "0" : s;
        }
    }
}
