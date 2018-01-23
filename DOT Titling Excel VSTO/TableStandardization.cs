using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace DOT_Titling_Excel_VSTO
{
    class TableStandardization
    {
        //Other Colors
        public static Color colorDullGreen = Color.FromArgb(169, 208, 142);
        public static Color colorDullBlue = Color.FromArgb(143, 172, 227);
        public static Color colorPinkAlt = Color.FromArgb(234, 174, 193);
        public static Color colorLightPink = Color.FromArgb(236, 204, 199);
        public static Color colorDarkBrown = Color.FromArgb(44, 36, 22);
        public static Color colorLightGreen = Color.FromArgb(198, 224, 180);
        public static Color colorLightDullGreen = Color.FromArgb(192, 207, 178);
        public static Color colorLightDullYellow = Color.FromArgb(255, 249, 186);
        public static Color colorLightDullBlue = Color.FromArgb(203, 233, 230);
        public static Color colorDarkGrey = Color.FromArgb(51, 51, 51);

        // Standard Colors
        public static XlRgbColor colorWhite = XlRgbColor.rgbGhostWhite;
        public static XlRgbColor colorBlack = XlRgbColor.rgbBlack;

        //Row Colors
        public static Color colorBugRow = colorLightDullYellow;
        public static Color colorDeletedRow = colorLightPink;
        public static Color colorReleaseRow = colorLightDullBlue;
        public static Color colorSprintRow = colorLightDullGreen;

        //Check Boxes (YesNo)
        // RED
        public static Color colorYesNoRed = colorPinkAlt;
        public static XlRgbColor colorYesNoRedFont = XlRgbColor.rgbDarkRed;
        // GREEN
        public static Color colorYesNoGreen = colorLightGreen;
        public static XlRgbColor colorYesNoGreenFont = XlRgbColor.rgbDarkGreen;
        // GOLD
        public static XlRgbColor colorYesNoGold = XlRgbColor.rgbGold;
        public static XlRgbColor colorYesNoGoldFont = XlRgbColor.rgbBrown;

        //Categories
        public static XlRgbColor colorCat1 = XlRgbColor.rgbDarkOliveGreen;
        public static XlRgbColor colorCat1Font = colorWhite;

        public static XlRgbColor colorCat2 = XlRgbColor.rgbForestGreen;
        public static XlRgbColor colorCat2Font = colorWhite;

        //Error Cells
        public static Color colorErrorCell = colorYesNoRed;
        public static XlRgbColor colorErrorCellFont = colorYesNoRedFont;
        public static XlRgbColor colorErrorCellBorder = XlRgbColor.rgbDimGrey; 

        public enum StandardizationType
        {
            Thorough = 1,
            Light = 2
        };

        public enum ColumnWidth
        {
            TextLong = 40,
            TextMedium = 20,
            TextShort = 15,
            TextTiny = 9,
            Number = 9,
            Dollar = 9,
            Decimal = 9,
            Percent = 9,
            Date = 12,
            Error = 7,
            YesNoGreen = 7,
            YesNoGold = 7,
            YesNoRed = 7,
            YesNo = 7,
            MidLong = 13,
            Release = 7,
            TicketType = 15,
            Hidden = 0,
            Default = 15
        };

        public static int GetColumnWidth(string ct)
        {
            try
            {
                if (ct == string.Empty || ct == null)
                ct = "Default";
                int cw = (int)((ColumnWidth)Enum.Parse(typeof(ColumnWidth), ct, true));
                return cw;
            }
            catch
            {
                MessageBox.Show("Not a recognized column type :" + ct);
                return (int)((ColumnWidth)Enum.Parse(typeof(ColumnWidth), "Default", true));
            }
        }

        public static void ExecuteCleanupTable(Excel.Application app, StandardizationType sType)
        {
            try
            {
                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    int column = 0;
                    Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);

                    string colType;
                    string columnHeader;

                    // Format each cell in the table header row and set column width
                    foreach (Range cell in headerRowRange.Cells)
                    {
                        column = cell.Column;
                        columnHeader = cell.Value;
                        colType = cell.Offset[-1, 0].Value;
                        cell.IndentLevel = 0;
                        if (colType == "TextLong")
                                cell.IndentLevel = 1;
                        cell.EntireColumn.ColumnWidth = GetColumnWidth(colType);
                    }

                    // Format the first row in the worksheet
                    Worksheet activeWorksheet = app.ActiveSheet;
                    Range r = activeWorksheet.get_Range("A1");
                    r.EntireRow.RowHeight = 40;

                    // Format the table header row
                    headerRowRange.EntireRow.RowHeight = 66;
                    headerRowRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    headerRowRange.VerticalAlignment = XlVAlign.xlVAlignTop;
                    headerRowRange.Font.Size = 9;
                    headerRowRange.VerticalAlignment = XlVAlign.xlVAlignTop;

                    // Format the table properties row
                    headerRowRange.Offset[-1, 0].Font.Size = 9;
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;

                    // Perform thorough standardization
                    if (sType == StandardizationType.Thorough)
                    {
                        ThoroughColumnCleanup(app, tableRangeName, headerRowRange);
                        ThoroughFooterCleanup(app, headerRowRange);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void ThoroughFooterCleanup(Excel.Application app, Range headerRowRange)
        {
            int headerRow = headerRowRange.Row;
            int footerRow = 0;
            int footerRowOffset = 0;
            string footerRangeName = SSUtils.GetSelectedTableFooter(app);
            if (footerRangeName != string.Empty)
            {
               
                Range footerRowRange = app.get_Range(footerRangeName, Type.Missing);
                footerRow = footerRowRange.Row;
                footerRowOffset = footerRow - headerRow;
                headerRowRange.Copy(Type.Missing);
                footerRowRange.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            }

            // Get the footer row and format it
            if (headerRow > 2)
            {
                Range propertiesRowRange = headerRowRange.Offset[-1, 0];
                if (propertiesRowRange != null)
                {
                    headerRowRange.Copy(Type.Missing);
                    propertiesRowRange.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
            }
        }

        private static void ThoroughColumnCleanup(Excel.Application app, string tableRangeName, Range headerRowRange)
        {
            //Conditional Formatting
            //https://stackoverflow.com/questions/11858529/deleting-a-conditionalformat
            Range tableRange = app.get_Range(tableRangeName, Type.Missing);
            if (tableRange != null)
            {
                tableRange.ClearFormats();
                int column;
                string columnHeader;
                string colType;
                int firstDataRow = headerRowRange.Row + 1;
                foreach (Range cell in headerRowRange.Cells)
                {
                    column = cell.Column;
                    columnHeader = cell.Value;
                    colType = cell.Offset[-1, 0].Value;
                    string columnNameRange = tableRangeName + '[' + columnHeader + ']';
                    Range columnRange = app.get_Range(columnNameRange, Type.Missing);
                    if (columnRange != null)
                    {
                        switch (colType)
                        {
                            case "Decimal":
                            case "Dollar":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                columnRange.NumberFormat = "#,##0.00";
                                break;
                            case "Number":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                break;
                            case "Percent":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                columnRange.NumberFormat = "0%";
                                break;
                            case "Date":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                columnRange.NumberFormat = "m/d/yyyy";
                                break;
                            case "YesNoGreen":
                                FormatYesNo(columnRange, colType);
                                break;
                            case "YesNoGold":
                                FormatYesNo(columnRange, colType);
                                break;
                            case "Error":
                            case "YesNoRed":
                            case "YesNo":
                                FormatYesNo(columnRange, colType);
                                break;
                            case "MidLong":
                                FormatMidLong(columnRange);
                                break;
                            case "Release":
                                FormatRelease(columnRange);
                                break;
                            case "TicketType":
                                FormatTicketType(columnRange);
                                break;
                            case "Hidden":
                                break;
                            default:
                                break;
                        }
                    }
                }
                FormatErrorColumns(app, tableRange, tableRangeName, firstDataRow);
                FormatRowsConditionally(app, tableRange, tableRangeName, firstDataRow);
            }            
        }

        private static void FormatErrorColumns(Excel.Application app, Range tableRange, string tableRangeName, int firstDataRow)
        {
            if (tableRangeName == "TicketData")
            {
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Summary", "Jira Summary" });

                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Epics Dont Match", new string[] { "Epic", "Jira Epic" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Releases Dont Match", new string[] { "Release", "Jira Release" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR No Backlog Area", new string[] { "Backlog Area" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Sprints Dont Match", new string[] { "Hufflepuff Sprint", "Jira Hufflepuff Sprint" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Dupe", new string[] { "Ticket ID" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR No Epic", new string[] { "Jira Epic ID", "Epic" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Points but To Do", new string[] { "Jira Status" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Current Sprint But No Points", new string[] { "Points" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Done No Sprint", new string[] { "Jira Hufflepuff Sprint" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Bug Not Categorized", new string[] { "DOT Jira ID" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Can be Deleted", new string[] { "Ticket Type" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Story Not Moving or Blocked", new string[] { "Days in Same Status", "Jira Status", "Jira Status (Last Changed)" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Need Reason for Blocker", new string[] { "Reason Blocked or Delayed"});
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Should be Assigned to Dev", new string[] { "Jira Status", "Role", "DOT Sprint" });

                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Workflow DOT Created", new string[] { "Backlog Area", "Date Submitted to DOT", "Date Approved by DOT", "Points" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Workflow DOT Written", new string[] { "Backlog Area", "Date Submitted to DOT", "Date Approved by DOT", "Points" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Workflow DOT Groomed", new string[] { "Backlog Area", "Date Submitted to DOT", "Date Approved by DOT", "Points" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Workflow DOT Submitted", new string[] { "Backlog Area", "Date Submitted to DOT", "Date Approved by DOT", "Points" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Workflow Ready for Development", new string[] { "Backlog Area", "Date Submitted to DOT", "Date Approved by DOT", "Points" });
            }

            if (tableRangeName == "DOTReleaseData")
            {
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Summary", "Jira Summary" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Epics Dont Match", new string[] { "Epic", "Jira Epic" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Dupe", new string[] { "Ticket ID" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR No Epic", new string[] { "Jira Epic ID", "Epic" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Points but To Do", new string[] { "Jira Status" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Done No Sprint", new string[] { "Jira Hufflepuff Sprint" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Bug Not Categorized", new string[] { "DOT Jira ID" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Can be Deleted", new string[] { "Ticket Type" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Multiple Releases", new string[] { "Jira Release" });
            }

            if (tableRangeName == "EpicData")
            {
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Jira Epic Summary", "Epic" });
                FormatErrorColumn(app, tableRangeName, firstDataRow, "ERR Points Dont Match", new string[] { "Jira Epic Points", "Estimate 2" });
            }
        }

        private static void FormatErrorColumn(Excel.Application app, string tableRangeName, int firstDataRow, string errField, string[] columns)
        {
            try
            {
                foreach (string column in columns)
                {
                    Range columnRange = app.get_Range(tableRangeName + "[" + column + "]", Type.Missing);
                    if (columnRange != null)
                    {
                        Range errorRange = app.get_Range(tableRangeName + "[" + errField + "]", Type.Missing);
                        if (errorRange != null)
                        {
                            string col = SSUtils.GetColumnName(columnRange.Column);
                            string errorCol = SSUtils.GetColumnName(errorRange.Column);
                            if (col != string.Empty)
                            {
                                string cond = "=$" + errorCol + firstDataRow + "=" + @"""x""";
                                FormatCondition fc = (FormatCondition)columnRange.FormatConditions.Add
                                    (XlFormatConditionType.xlExpression, Type.Missing, cond, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                fc.Interior.Color = colorErrorCell;
                                fc.Font.Color = colorErrorCellFont;
                                fc.Font.Bold = true;
                            }
                        }
                    }
                }
            }
            catch
            {
                //MessageBox.Show("Error :" + ex);
                //return string.Empty;
            }
        }

        private static void FormatRowsConditionally(Excel.Application app, Range tableRange, string tableRangeName, int firstDataRow)
        {
            if (tableRangeName == "TicketData" || tableRangeName == "DOTReleaseData")
            {
                //Ticket Type Column
                Range ticketTypeColumnRange = app.get_Range(tableRangeName + "[Ticket Type]", Type.Missing);
                string ticketTypeColumn = SSUtils.GetColumnName(ticketTypeColumnRange.Column);

                //Bug Row
                string condBug = "=$" + ticketTypeColumn + firstDataRow + "=" + @"""Software Bug""";
                FormatCondition fcBug = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, condBug, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcBug.Interior.Color = colorBugRow;

                //Deleted Ticket Row
                string conDeleted = "=$" + ticketTypeColumn + firstDataRow + "=" + @"""{DELETED}""";
                FormatCondition fcDeleted = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, conDeleted, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcDeleted.Interior.Color = colorDeletedRow;
            }
            else
            {
                if (tableRangeName == "SprintData" || tableRangeName == "ReleaseData" || tableRangeName == "EpicData")
                {
                    //Release Column
                    Range releaseColumnRange = app.get_Range(tableRangeName + "[R]", Type.Missing);
                    string releaseColumn = SSUtils.GetColumnName(releaseColumnRange.Column);

                    //Current Sprint Row
                    if (tableRangeName == "SprintData")
                    {
                        //Sprint Column
                        Range sprintColumnRange = app.get_Range(tableRangeName + "[Sprint]", Type.Missing);
                        string sprintColumn = SSUtils.GetColumnName(sprintColumnRange.Column);

                        string conSelected = "=$" + sprintColumn + firstDataRow + "=CurrentSprint";
                        FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                            (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        fcSelected.Interior.Color = colorSprintRow;

                        string conSelected1 = "=AND($" + releaseColumn + firstDataRow + "=CurrentRelease,$" + sprintColumn + firstDataRow + "<>CurrentSprint)";
                        FormatCondition fcSelected1 = (FormatCondition)tableRange.FormatConditions.Add
                            (XlFormatConditionType.xlExpression, Type.Missing, conSelected1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        fcSelected1.Interior.Color = colorReleaseRow;
                    }

                    //Current Release Row
                    if (tableRangeName == "ReleaseData")
                    {
                        string conSelected = "=$" + releaseColumn + firstDataRow + "=CurrentRelease";
                        FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                            (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        fcSelected.Interior.Color = colorReleaseRow;
                    }

                    //Current Release Row
                    if (tableRangeName == "EpicData")
                    {
                        string conSelected = "=$" + releaseColumn + firstDataRow + "=CurrentRelease";
                        FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                            (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        fcSelected.Interior.Color = colorReleaseRow;
                    }
                }
            }
        }

        private static void FormatTicketType(Range columnRange)
        {
            // Ticket
            FormatCondition cStory = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Story", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cStory.Interior.Color = colorLightDullGreen;
            cStory.Font.Color = XlRgbColor.rgbDarkGreen;
            cStory.Font.Color = colorBlack;

            // Software Bug
            FormatCondition cBug = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Software Bug", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cBug.Interior.Color = colorBugRow;
            //cBug.Font.Color = colorDarkBrown;
            cBug.Font.Color = colorBlack;

            // {DELETED}
            FormatCondition cDeleted = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "{DELETED}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cDeleted.Interior.Color = colorDeletedRow;
            //cDeleted.Font.Color = XlRgbColor.rgbDarkRed;
            cDeleted.Font.Color = colorBlack;
        }

        private static void FormatMidLong(Range columnRange)
        {
            // MID
            FormatCondition cMID = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Mid", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cMID.Interior.Color = colorCat1;
            cMID.Font.Color = colorCat1Font;

            // LONG
            FormatCondition cLong = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Long", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cLong.Interior.Color = colorCat2;
            cLong.Font.Color = colorCat2Font;

            // Phase 2
            FormatCondition cPhase2 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Phase 2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cPhase2.Interior.Color = colorBlack;
            cPhase2.Font.Color = colorWhite;

            // Out of Scope
            FormatCondition cOOS = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Out of Scope", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cOOS.Interior.Color = colorDarkGrey;
            cOOS.Font.Color = colorWhite;

            columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            columnRange.Font.Bold = true;
        }

        private static void FormatYesNo(Range columnRange, string colType)
        {
            // RGB Colors
            // http://www.flounder.com/csharp_color_table.htm

            FormatCondition condition =
                   (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "x",
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing);

            switch (colType)
            {
                case "YesNoRed":
                    condition.Interior.Color = colorYesNoRed;
                    condition.Font.Color = colorYesNoRedFont;
                    break;
                case "YesNoGreen":
                    condition.Interior.Color = colorYesNoGreen;
                    condition.Font.Color = colorYesNoGreenFont;
                    break;
                case "YesNoGold":
                    condition.Interior.Color = colorYesNoGold;
                    condition.Font.Color = colorYesNoGoldFont;
                    break;
                default:
                    condition.Interior.Color = colorYesNoRed;
                    condition.Font.Color = colorYesNoRedFont;
                    break;
            }
            condition.Font.Bold = true;
            columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private static void FormatRelease(Range columnRange)
        {
            // From http://colorbrewer2.org
            // Convert Hex to RGB: http://www.javascripter.net/faq/hextorgb.htm
            // https://blog.graphiq.com/finding-the-right-color-palettes-for-data-visualizations-fcd4e707a283

            Color c1 = Color.FromArgb(45, 15, 65);
            Color c2 = Color.FromArgb(61, 20, 89);
            Color c3 = Color.FromArgb(77, 26, 112);
            Color c4 = Color.FromArgb(94, 31, 136);
            Color c5 = Color.FromArgb(116, 39, 150);
            Color c6 = Color.FromArgb(151, 52, 144);
            Color c7 = Color.FromArgb(184, 66, 140);
            Color c8 = Color.FromArgb(219, 80, 135);
            Color c9 = Color.FromArgb(233, 106, 141);
            Color c10 = Color.FromArgb(238, 139, 151);
            Color c11 = Color.FromArgb(243, 172, 162);
            Color c12 = Color.FromArgb(249, 205, 172);
            Color c13 = Color.FromArgb(67, 147, 195); // TO DO
            Color c14 = Color.FromArgb(33, 102, 172); // TO DO
            Color c15 = Color.FromArgb(5, 48, 97);    // TO DO

            // R1
            FormatCondition conditionR1 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            conditionR1.Interior.Color = c1;
            conditionR1.Font.Color = colorWhite;

            // R2
            FormatCondition conditionR2 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR2.Interior.Color = c2;
            conditionR2.Font.Color = colorWhite;

            // R3
            FormatCondition conditionR3 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R3", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR3.Interior.Color = c3;
            conditionR3.Font.Color = colorWhite;

            // R4
            FormatCondition conditionR4 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R4", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR4.Interior.Color = c4;
            conditionR4.Font.Color = colorWhite;

            // R5
            FormatCondition conditionR5 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R5", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR5.Interior.Color = c5;
            conditionR5.Font.Color = colorWhite;

            // R6
            FormatCondition conditionR6 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R6", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR6.Interior.Color = c6;
            conditionR6.Font.Color = colorWhite;

            // R7
            FormatCondition conditionR7 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R7", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR7.Interior.Color = c7;
            conditionR7.Font.Color = colorWhite;

            // R8
            FormatCondition conditionR8 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R8", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR8.Interior.Color = c8;
            conditionR8.Font.Color = colorWhite;

            // R9
            FormatCondition conditionR9 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R9", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR9.Interior.Color = c9;
            conditionR9.Font.Color = colorWhite;

            // R10
            FormatCondition conditionR10 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R10", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR10.Interior.Color = c10;
            conditionR10.Font.Color = colorWhite;

            // R11
            FormatCondition conditionR11 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R11", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR11.Interior.Color = c11;
            conditionR11.Font.Color = colorWhite;

            // R12
            FormatCondition conditionR12 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R12", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR12.Interior.Color = c12;
            conditionR12.Font.Color = colorWhite;

            // R13
            FormatCondition conditionR13 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R13", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR13.Interior.Color = c13;
            conditionR13.Font.Color = colorWhite;

            // R14
            FormatCondition conditionR14 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R14", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR14.Interior.Color = c14;
            conditionR14.Font.Color = colorWhite;

            // R15
            FormatCondition conditionR15 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R15", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR15.Interior.Color = c15;
            conditionR15.Font.Color = colorWhite;

            // R98
            FormatCondition conditionR98 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R98", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR98.Interior.Color = colorBlack;
            conditionR98.Font.Color = colorWhite;

            // R99
            FormatCondition conditionR99 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R99", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR99.Interior.Color = colorDarkGrey;
            conditionR99.Font.Color = colorWhite;

            columnRange.Font.Bold = true;
            columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void ExecuteShowHidePropertiesRow(Excel.Application app)
        {
            try
            {
                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    int headerRow = 0;
                    int propertiesRow = 0;

                    Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    headerRow = headerRowRange.Row;
                    propertiesRow = headerRowRange.Row - 1;

                    if (propertiesRow > 0)
                    {
                        Worksheet ws = app.ActiveSheet;
                        bool hidden = ws.Rows[propertiesRow].EntireRow.Height == 0;
                        headerRowRange.EntireRow.Offset[-1, 0].Hidden = !hidden;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
