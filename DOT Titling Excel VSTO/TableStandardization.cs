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
        public static Color colorDarkBrown = Color.FromArgb(44, 36, 22);

        // Standard Colors
        public static XlRgbColor colorWhite = XlRgbColor.rgbGhostWhite;
        public static XlRgbColor colorBlack = XlRgbColor.rgbBlack;
        public static XlRgbColor colorDarkGrey = XlRgbColor.rgbDarkSlateGray; 

        //Row Colors
        public static XlRgbColor colorBugRow = XlRgbColor.rgbPaleGoldenrod;
        public static XlRgbColor colorDeletedRow = XlRgbColor.rgbPink;

        //Check Boxes (YesNo)
        // RED
        public static Color colorYesNoRed = colorPinkAlt;
        public static XlRgbColor colorYesNoRedFont = XlRgbColor.rgbDarkRed;
        // GREEN
        public static XlRgbColor colorYesNoGreen = XlRgbColor.rgbLightGreen;
        public static XlRgbColor colorYesNoGreenFont = XlRgbColor.rgbDarkGreen;
        // GOLD
        public static XlRgbColor colorYesNoGold = XlRgbColor.rgbGold;
        public static XlRgbColor colorYesNoGoldFont = XlRgbColor.rgbBrown;

        //Categories
        public static Color colorCat1 = colorDullGreen;
        public static XlRgbColor colorCat1Font = XlRgbColor.rgbDarkGreen;

        public static Color colorCat2 = colorDullBlue;
        public static XlRgbColor colorCat2Font = XlRgbColor.rgbDarkSlateBlue;

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
            Percent = 9,
            Date = 10,
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
            //ColumnWidth cw = (ColumnWidth)Enum.Parse(typeof(ColumnWidth), colour, true); 
            if (ct == string.Empty || ct == null)
                ct = "Default";
            int cw = (int)((ColumnWidth)Enum.Parse(typeof(ColumnWidth), ct, true));
            return cw;
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
                            case "Number":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                break;
                            case "Percent":
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                columnRange.NumberFormat = "0%";
                                break;
                            case "Date":
                                cell.EntireColumn.ColumnWidth = 10;
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
                FormatRows(tableRange, tableRangeName);
            }            
        }

        private static void FormatRows(Range tableRange, string tableRangeName)
        {
            //Bug Row
            string condBug = "=$A4=" + @"""Software Bug""";
            FormatCondition fcBug = (FormatCondition)tableRange.FormatConditions.Add
                (XlFormatConditionType.xlExpression, Type.Missing, condBug, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            fcBug.Interior.Color = colorBugRow;

            //Deleted Ticket Row
            string conDeleted = "=$A4=" + @"""{DELETED}""";
            FormatCondition fcDeleted = (FormatCondition)tableRange.FormatConditions.Add
                (XlFormatConditionType.xlExpression, Type.Missing, conDeleted, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            fcDeleted.Interior.Color = colorDeletedRow;

            //Current Sprint Row
            if (tableRangeName == "SprintData")
            {
                string conSelected = "=$A5=CurrentSprint";
                FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcSelected.Interior.Color = colorDullGreen;

                string conSelected1 = "=AND($L5=CurrentRelease,$A5<>CurrentSprint)";
                FormatCondition fcSelected1 = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, conSelected1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcSelected1.Interior.Color = colorDullBlue;
            }

            //Current Release Row
            if (tableRangeName == "ReleaseData")
            {
                string conSelected = "=$B5=CurrentRelease";
                FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcSelected.Interior.Color = colorDullBlue;
            }

            //Current Release Row
            if (tableRangeName == "EpicData")
            {
                string conSelected = "=$F4=CurrentRelease";
                FormatCondition fcSelected = (FormatCondition)tableRange.FormatConditions.Add
                    (XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                fcSelected.Interior.Color = colorDullBlue;
            }
        }

        private static void FormatRelease(Range columnRange)
        {
            // From http://colorbrewer2.org
            Color c1 = Color.FromArgb(158,1,66);
            Color c2 = Color.FromArgb(213,62,79);
            Color c3 = Color.FromArgb(244,109,67);
            Color c4 = Color.FromArgb(253,174,97);
            Color c5 = Color.FromArgb(254,224,139);
            Color c6 = Color.FromArgb(255,255,191);
            Color c7 = Color.FromArgb(230,245,152);
            Color c8 = Color.FromArgb(171,221,164);
            Color c9 = Color.FromArgb(102,194,165);
            Color c10 = Color.FromArgb(50,136,189);
            Color c11 = Color.FromArgb(94,79,162);
            Color c12 = Color.FromArgb(146,197,222);
            Color c13 = Color.FromArgb(67,147,195);
            Color c14 = Color.FromArgb(33,102,172);
            Color c15 = Color.FromArgb(5,48,97);

            // R1
            FormatCondition conditionR1 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R1", Type.Missing, Type.Missing,  Type.Missing, Type.Missing, Type.Missing);
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
            conditionR5.Font.Color = colorBlack;

            // R6
            FormatCondition conditionR6 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R6", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR6.Interior.Color = c6;
            conditionR6.Font.Color = colorBlack;

            // R7
            FormatCondition conditionR7 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R7", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR7.Interior.Color = c7;
            conditionR7.Font.Color = colorBlack;

            // R8
            FormatCondition conditionR8 = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "R8", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR8.Interior.Color = c8;
            conditionR8.Font.Color = colorBlack;

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

        private static void FormatTicketType(Range columnRange)
        {
            Color dullGreen = Color.FromArgb(169, 208, 142);

            // Ticket
            FormatCondition cStory = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Story", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cStory.Interior.Color = dullGreen;
            cStory.Font.Color = XlRgbColor.rgbDarkGreen;

            // Software Bug
            FormatCondition cBug = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "Software Bug", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cBug.Interior.Color = colorBugRow;
            cBug.Font.Color = colorDarkBrown;

            // {DELETED}
            FormatCondition cDeleted = (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                   XlFormatConditionOperator.xlEqual, "{DELETED}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cDeleted.Interior.Color = colorDeletedRow;
            cDeleted.Font.Color = XlRgbColor.rgbDarkRed;

            columnRange.Font.Bold = true;
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
