using System;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;

namespace DOT_Titling_Excel_VSTO
{
    class TableStandardization
    {
        public enum ColumnWidth
        {
            TextLong = 40,
            TextExtraLong = 110,
            TextMedium = 20,
            TextShort = 15,
            TextTiny = 9,
            Priority = 20,
            Number = 9,
            NumberTiny = 7,
            BaseballAvg = 7,
            Dollar = 9,
            Decimal = 9,
            Percent = 9,
            Percent1 = 9,
            Date = 12,
            Time = 12,
            Error = 7,
            YesNoGreen = 7,
            YesNoGold = 7,
            YesNoRed = 7,
            YesNo = 7,
            YesOrNo = 7,
            MidLong = 13,
            R = 7,
            IssueType = 20,
            ProjectKey = 15,
            Hidden = 0,
            Default = 15,
            Special = 10
        };

        public enum ColumnExceptionType
        {
            Error,
            Warning,
            Workflow
        }

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
        public static Color colorSFR = Color.FromArgb(172, 185, 202);
        public static Color colorTask = Color.FromArgb(248, 203, 173);

        // Standard Colors
        public static Excel.XlRgbColor colorWhite = Excel.XlRgbColor.rgbGhostWhite;
        public static Excel.XlRgbColor colorBlack = Excel.XlRgbColor.rgbBlack;
        public static Excel.XlRgbColor colorNavy = Excel.XlRgbColor.rgbNavyBlue;

        //Row Colors
        public static Color colorBugRow = colorLightDullYellow;
        public static Color colorDeletedRow = colorLightPink;
        public static Color colorReleaseRow = colorLightDullBlue;
        public static Color colorSprintRow = colorLightDullGreen;

        //Check Boxes (YesNo)
        // RED
        public static Color colorYesNoRed = colorPinkAlt;
        public static Excel.XlRgbColor colorYesNoRedFont = Excel.XlRgbColor.rgbDarkRed;
        // GREEN
        public static Color colorYesNoGreen = colorLightGreen;
        public static Excel.XlRgbColor colorYesNoGreenFont = Excel.XlRgbColor.rgbDarkGreen;
        // GOLD
        public static Excel.XlRgbColor colorYesNoGold = Excel.XlRgbColor.rgbGold;
        public static Excel.XlRgbColor colorYesNoGoldFont = Excel.XlRgbColor.rgbBrown;

        //Categories
        public static Excel.XlRgbColor colorCat1 = Excel.XlRgbColor.rgbDarkOliveGreen;
        public static Excel.XlRgbColor colorCat1Font = colorWhite;

        public static Excel.XlRgbColor colorCat2 = Excel.XlRgbColor.rgbForestGreen;
        public static Excel.XlRgbColor colorCat2Font = colorWhite;

        //Error Cells
        public static Color colorErrorCell = colorYesNoRed;
        public static Excel.XlRgbColor colorErrorCellFont = colorYesNoRedFont;
        public static Excel.XlRgbColor colorErrorCellBorder = Excel.XlRgbColor.rgbDimGrey;

        //Warning Cells
        public static Excel.XlRgbColor colorWarningCell = colorYesNoGold;
        public static Excel.XlRgbColor colorWarningCellFont = colorYesNoGoldFont;
        public static Excel.XlRgbColor colorWarningCellBorder = Excel.XlRgbColor.rgbDimGrey;

        //Workflow Cells
        public static Color colorWorkflowCell = colorLightGreen;
        public static Excel.XlRgbColor colorWorkflowCellFont = colorYesNoGreenFont;
        public static Excel.XlRgbColor colorWorkflowCellBorder = Excel.XlRgbColor.rgbDimGrey;

        public enum StandardizationType
        {
            Thorough = 1,
            Light = 2
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

        public async static Task<bool> Execute(Excel.Application app, StandardizationType sType)
        {
            try
            {
                bool success;
                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    int column = 0;
                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);

                    string colType;
                    string columnHeader;

                    // Format each cell in the table header row and set column width
                    foreach (Excel.Range cell in headerRowRange.Cells)
                    {
                        column = cell.Column;
                        columnHeader = cell.Value;
                        colType = cell.Offset[-1, 0].Value;
                        cell.IndentLevel = 0;
                        if (colType == "TextLong" || colType == "TextExtraLong")
                                cell.IndentLevel = 1;
                        cell.EntireColumn.ColumnWidth = GetColumnWidth(colType);
                    }

                    // Format the first row in the worksheet
                    Excel.Worksheet ws = app.ActiveSheet;
                    Excel.Range r = ws.get_Range("A1");
                    r.EntireRow.RowHeight = 40;

                    // Format the table header row
                    headerRowRange.EntireRow.RowHeight = 66;
                    headerRowRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    headerRowRange.Font.Size = 9;
                    headerRowRange.Font.Name = "Calibri Light";
                    headerRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                    // Format the table properties row
                    headerRowRange.Offset[-1, 0].Font.Size = 9;
                    headerRowRange.Offset[-1, 0].Font.Name = "Calibri Light";
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;

                    // Perform thorough standardization
                    if (sType == StandardizationType.Thorough)
                    {
                        success = await ThoroughColumnCleanup(app, tableRangeName, headerRowRange);
                        success = await ThoroughFooterCleanup(app, headerRowRange);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private async static Task<bool> ThoroughFooterCleanup(Excel.Application app, Excel.Range headerRowRange)
        {
            try
            {
                int headerRow = headerRowRange.Row;
                int footerRow = 0;
                int footerRowOffset = 0;
                string footerRangeName = SSUtils.GetSelectedTableFooter(app);
                if (footerRangeName != string.Empty)
                {

                    Excel.Range footerRowRange = app.get_Range(footerRangeName, Type.Missing);
                    footerRow = footerRowRange.Row;
                    footerRowOffset = footerRow - headerRow;
                    headerRowRange.Copy(Type.Missing);
                    // PWH TO DO - NEED TO FIX THIS
                    //footerRowRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    //footerRowRange.Font.Name = "Calibri Light";
                }

                // Get the footer row and format it
                if (headerRow > 2)
                {
                    Excel.Range propertiesRowRange = headerRowRange.Offset[-1, 0];
                    if (propertiesRowRange != null)
                    {
                        headerRowRange.Copy(Type.Missing);
                        propertiesRowRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private async static Task<bool> ThoroughColumnCleanup(Excel.Application app, string tableRangeName, Excel.Range headerRowRange)
        {
            try
            {                
                //Conditional Formatting
                //https://stackoverflow.com/questions/11858529/deleting-a-conditionalformat
                Excel.Range tableRange = app.get_Range(tableRangeName, Type.Missing);
                if (tableRange != null)
                {
                    tableRange.ClearFormats();
                    int column;
                    string columnHeader;
                    string colType;
                    int firstDataRow = headerRowRange.Row + 1;
                    foreach (Excel.Range cell in headerRowRange.Cells)
                    {
                        column = cell.Column;
                        columnHeader = cell.Value;
                        colType = cell.Offset[-1, 0].Value;
                        string columnNameRange = tableRangeName + '[' + columnHeader + ']';
                        Excel.Range columnRange = app.get_Range(columnNameRange, Type.Missing);
                        if (columnRange != null)
                        {
                            switch (colType)
                            {
                                case "Decimal":
                                case "Dollar":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    columnRange.NumberFormat = "#,##0.00";
                                    break;
                                case "BaseballAvg":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    columnRange.NumberFormat = "#,##0.000";
                                    break;
                                case "NumberTiny":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    break;
                                case "Number":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    break;
                                case "Percent":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    columnRange.NumberFormat = "0%";
                                    break;
                                case "Percent1":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    columnRange.NumberFormat = "0.0%";
                                    break;
                                case "Date":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    columnRange.NumberFormat = "m/d/yyyy";
                                    break;
                                case "Time":
                                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    columnRange.NumberFormat = "h:mm AM/PM";
                                    break;
                                case "YesOrNo":
                                    FormatYesOrNo(columnRange);
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
                                case "Special":
                                    FormatSpecial(columnRange);
                                    break;
                                case "R":
                                    FormatR(columnRange);
                                    break;
                                case "IssueType":
                                    FormatIssueType(columnRange);
                                    break;
                                case "Priority":
                                    FormatPriority(columnRange);
                                    break;
                                case "Hidden":
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    bool success;
                    success = await FormatExceptionColumns(app, tableRange, tableRangeName, firstDataRow);
                    success = await FormatRowsConditionally(app, tableRange, tableRangeName, firstDataRow);
                    tableRange.Font.Name = "Calibri Light";
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private async static Task<bool> FormatExceptionColumns(Excel.Application app, Excel.Range tableRange, string tableRangeName, int firstDataRow)
        {
            bool success;
            if (tableRangeName == "IssueData")
            {
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Summary (Local)", "Summary" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Epics Dont Match", new string[] { "Epic (Local)", "Epic", "Epic Link" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Releases Dont Match", new string[] { "Release (Local)", "Fix Version" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR No Sprint", new string[] { "Sprint Number (Local)", "Sprint Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Sprints Dont Match", new string[] { "Sprint Number (Local)", "Sprint Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Dupe", new string[] { "Issue ID" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR No Epic", new string[] { "Epic (Local)", "Epic", "Epic Link" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Points but To Do", new string[] { "Story Points", "Status" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Current Sprint But No Points", new string[] { "Story Points" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Done No Sprint", new string[] { "Sprint Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Bug Not Categorized", new string[] { "DOT Jira ID" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Can be Deleted", new string[] { "Issue Type" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Should be Assigned to Dev", new string[] { "Status", "Role", "Sprint Number (Local)" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Story from Previous Release should be done", new string[] { "WIN Release", "Epic Release Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Warning, tableRangeName, firstDataRow, "WARN Story Not Moving or Blocked", new string[] { "Days in Same Status", "Status", "Status (Last Changed)" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Warning, tableRangeName, firstDataRow, "WARN Need Reason for Blocker", new string[] { "Reason Blocked or Delayed" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Warning, tableRangeName, firstDataRow, "WARN Check Bypass Approval", new string[] { "Bypass Approval", "Date Submitted to DOT", "Date Approved by DOT" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Workflow, tableRangeName, firstDataRow, "WFLOW Created", new string[] { "Sprint", "Story Points" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Workflow, tableRangeName, firstDataRow, "WFLOW Written", new string[] { "Sprint", "Story Points" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Workflow, tableRangeName, firstDataRow, "WFLOW Ready", new string[] { "Sprint", "Story Points" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Workflow, tableRangeName, firstDataRow, "WFLOW Bug Bucket", new string[] { "Sprint", "Status", "Sprint Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Workflow, tableRangeName, firstDataRow, "WFLOW Bug Bucket but Not a Bug", new string[] { "Sprint" });
            }

            if (tableRangeName == "DOTReleaseData")
            {
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Summary (Local)", "Summary" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Epics Dont Match", new string[] { "Epic (Local)", "Epic", "Epic Link" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Dupe", new string[] { "Issue ID" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR No Epic", new string[] { "Epic(Local)", "Epic", "Epic Link" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Points but To Do", new string[] { "Status" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Done No Sprint", new string[] { "Sprint Number" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Bug Not Categorized", new string[] { "DOT Jira ID" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Can be Deleted", new string[] { "Issue Type" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Multiple Releases", new string[] { "Fix Version" });
            }

            if (tableRangeName == "EpicData")
            {
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Summaries Dont Match", new string[] { "Epic", "Summary" });
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Points Dont Match", new string[] { "Story Points", "Estimate 4" });
            }

            if (tableRangeName == "ProjectsData")
            {
                success = await FormatExceptionColumn(app, ColumnExceptionType.Error, tableRangeName, firstDataRow, "ERR Project Names Dont Match", new string[] { "Project Name (Local)", "Project Name" });
            }
            return true;
        }

        private async static Task<bool> FormatExceptionColumn(Excel.Application app, ColumnExceptionType cType, string tableRangeName, int firstDataRow, string errField, string[] columns)
        {
            try
            {
                foreach (string column in columns)
                {
                    Excel.Range columnRange = app.get_Range(tableRangeName + "[" + column + "]", Type.Missing);
                    if (columnRange != null)
                    {
                        Excel.Range errorRange = app.get_Range(tableRangeName + "[" + errField + "]", Type.Missing);
                        if (errorRange != null)
                        {
                            string col = SSUtils.GetColumnName(columnRange.Column);
                            string errorCol = SSUtils.GetColumnName(errorRange.Column);
                            if (col != string.Empty)
                            {
                                string cond = "=$" + errorCol + firstDataRow + "=" + @"""x""";
                                Excel.FormatCondition fc = (Excel.FormatCondition)columnRange.FormatConditions.Add
                                    (Excel.XlFormatConditionType.xlExpression, Type.Missing, cond, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                fc.Font.Bold = true;
                                switch (cType)
                                {
                                    case ColumnExceptionType.Error:
                                        fc.Interior.Color = colorErrorCell;
                                        fc.Font.Color = colorErrorCellFont;
                                        break;
                                    case ColumnExceptionType.Warning:
                                        fc.Interior.Color = colorWarningCell;
                                        fc.Font.Color = colorWarningCellFont;
                                        break;
                                    case ColumnExceptionType.Workflow:
                                        fc.Interior.Color = colorWorkflowCell;
                                        fc.Font.Color = colorWorkflowCellFont;
                                        break;
                                    default:
                                        fc.Interior.Color = colorErrorCell;
                                        fc.Font.Color = colorErrorCellFont;
                                        break;
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch
            {
                //MessageBox.Show("Error :" + ex);
                //return string.Empty;
                return false;
            }
        }

        private async static Task<bool> FormatRowsConditionally(Excel.Application app, Excel.Range tableRange, string tableRangeName, int firstDataRow)
        {
            try
            {
                if (tableRangeName == "IssueData" || tableRangeName == "DOTReleaseData")
                {
                    //Issue Type Column
                    Excel.Range issueTypeColumnRange = app.get_Range(tableRangeName + "[Issue Type]", Type.Missing);
                    string issueTypeColumn = SSUtils.GetColumnName(issueTypeColumnRange.Column);

                    //Software Bug Row
                    string condBug = "=$" + issueTypeColumn + firstDataRow + "=" + @"""Software Bug""";
                    Excel.FormatCondition fcBug = (Excel.FormatCondition)tableRange.FormatConditions.Add
                        (Excel.XlFormatConditionType.xlExpression, Type.Missing, condBug, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fcBug.Interior.Color = colorBugRow;

                    //Bug Row
                    string condBug1 = "=$" + issueTypeColumn + firstDataRow + "=" + @"""Bug""";
                    Excel.FormatCondition fcBug1 = (Excel.FormatCondition)tableRange.FormatConditions.Add
                        (Excel.XlFormatConditionType.xlExpression, Type.Missing, condBug1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fcBug1.Interior.Color = colorBugRow;

                    //Deleted Issue Row
                    string conDeleted = "=$" + issueTypeColumn + firstDataRow + "=" + @"""{DELETED}""";
                    Excel.FormatCondition fcDeleted = (Excel.FormatCondition)tableRange.FormatConditions.Add
                        (Excel.XlFormatConditionType.xlExpression, Type.Missing, conDeleted, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fcDeleted.Interior.Color = colorDeletedRow;
                }
                else
                {
                    if (tableRangeName == "SprintData" || tableRangeName == "ReleaseData" || tableRangeName == "EpicData")
                    {
                        //Release Column
                        Excel.Range releaseColumnRange = app.get_Range(tableRangeName + "[R]", Type.Missing);
                        string releaseColumn = SSUtils.GetColumnName(releaseColumnRange.Column);

                        //Current Sprint Row
                        if (tableRangeName == "SprintData")
                        {
                            //Sprint Column
                            Excel.Range sprintColumnRange = app.get_Range(tableRangeName + "[Sprint]", Type.Missing);
                            string sprintColumn = SSUtils.GetColumnName(sprintColumnRange.Column);

                            string conSelected = "=$" + sprintColumn + firstDataRow + "=CurrentSprint";
                            Excel.FormatCondition fcSelected = (Excel.FormatCondition)tableRange.FormatConditions.Add
                                (Excel.XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            fcSelected.Interior.Color = colorSprintRow;

                            string conSelected1 = "=AND($" + releaseColumn + firstDataRow + "=CurrentRelease,$" + sprintColumn + firstDataRow + "<>CurrentSprint)";
                            Excel.FormatCondition fcSelected1 = (Excel.FormatCondition)tableRange.FormatConditions.Add
                                (Excel.XlFormatConditionType.xlExpression, Type.Missing, conSelected1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            fcSelected1.Interior.Color = colorReleaseRow;
                        }

                        //Current Release Row
                        if (tableRangeName == "ReleaseData")
                        {
                            string conSelected = "=$" + releaseColumn + firstDataRow + "=CurrentRelease";
                            Excel.FormatCondition fcSelected = (Excel.FormatCondition)tableRange.FormatConditions.Add
                                (Excel.XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            fcSelected.Interior.Color = colorReleaseRow;
                        }

                        //Current Release Row
                        if (tableRangeName == "EpicData")
                        {
                            string conSelected = "=$" + releaseColumn + firstDataRow + "=CurrentRelease";
                            Excel.FormatCondition fcSelected = (Excel.FormatCondition)tableRange.FormatConditions.Add
                                (Excel.XlFormatConditionType.xlExpression, Type.Missing, conSelected, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            fcSelected.Interior.Color = colorReleaseRow;
                        }
                    }
                }
                return true;
            }
            catch
            {
                //MessageBox.Show("Error :" + ex);
                //return string.Empty;
                return false;
            }
        }

        private static void FormatIssueType(Excel.Range columnRange)
        {
            // Issue
            Excel.FormatCondition cStory = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Story", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cStory.Interior.Color = colorLightDullGreen;
            cStory.Font.Color = Excel.XlRgbColor.rgbDarkGreen;
            cStory.Font.Color = colorBlack;

            // Software Bug
            Excel.FormatCondition cBug = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Software Bug", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cBug.Interior.Color = colorBugRow;
            //cBug.Font.Color = colorDarkBrown;
            cBug.Font.Color = colorBlack;

            // Bug
            Excel.FormatCondition cBug1 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Bug", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cBug1.Interior.Color = colorBugRow;
            //cBug.Font.Color = colorDarkBrown;
            cBug1.Font.Color = colorBlack;


            // Task 
            Excel.FormatCondition cTask = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Task", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cTask.Interior.Color = colorTask;
            cTask.Font.Color = colorBlack;

            //Software Feature Request
            Excel.FormatCondition cSFR = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Software Feature Request", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cSFR.Interior.Color = colorSFR;
            cSFR.Font.Color = colorBlack;

            // {DELETED}
            Excel.FormatCondition cDeleted = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "{DELETED}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cDeleted.Interior.Color = colorDeletedRow;
            //cDeleted.Font.Color = Excel.XlRgbColor.rgbDarkRed;
            cDeleted.Font.Color = colorBlack;
        }

        private static void FormatPriority(Excel.Range columnRange)
        {
            // 1. Critical
            Excel.FormatCondition cStory = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "1. Critical", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cStory.Interior.Color = colorYesNoRed;
            cStory.Font.Color = colorYesNoRedFont;

            // 2. High
            Excel.FormatCondition cBug = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "2. High", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cBug.Interior.Color = colorYesNoGold;
            cBug.Font.Color = colorYesNoGoldFont;
        }

        private static void FormatMidLong(Excel.Range columnRange)
        {
            // MID
            Excel.FormatCondition cMID = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Mid", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cMID.Interior.Color = colorCat1;
            cMID.Font.Color = colorCat1Font;

            // LONG
            Excel.FormatCondition cLong = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Long", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cLong.Interior.Color = colorCat2;
            cLong.Font.Color = colorCat2Font;

            // Phase 2
            Excel.FormatCondition cPhase2 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Phase 2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cPhase2.Interior.Color = colorBlack;
            cPhase2.Font.Color = colorWhite;

            // Out of Scope
            Excel.FormatCondition cOOS = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Out of Scope", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cOOS.Interior.Color = colorDarkGrey;
            cOOS.Font.Color = colorWhite;

            columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            columnRange.Font.Bold = true;
        }


        private static void FormatSpecial(Excel.Range columnRange)
        {
            // WON
            Excel.FormatCondition cWON = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Won", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            cWON.Interior.Color = colorCat1;
            cWON.Font.Color = colorCat1Font;

            // LOST
            Excel.FormatCondition cLOST = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Lost", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cLOST.Interior.Color = colorYesNoRedFont;
            cLOST.Font.Color = colorCat1Font;


            // HOME
            Excel.FormatCondition cHOME = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Home", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cHOME.Interior.Color = colorNavy;
            cHOME.Font.Color = colorCat1Font;

            // AWAY
            Excel.FormatCondition cAWAY = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Away", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cAWAY.Interior.Color = colorDarkGrey;
            cAWAY.Font.Color = colorCat1Font;


            columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            columnRange.Font.Bold = true;
        }


        private static void FormatYesNo(Excel.Range columnRange, string colType)
        {
            // RGB Colors
            // http://www.flounder.com/csharp_color_table.htm

            Excel.FormatCondition condition =
                   (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "x",
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
            columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }


        private static void FormatYesOrNo(Excel.Range columnRange)
        {
            // RGB Colors
            // http://www.flounder.com/csharp_color_table.htm

            Excel.FormatCondition conditionYes =
                   (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "Yes",
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing);

            Excel.FormatCondition conditionNo =
                   (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "No",
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing);

                    conditionNo.Interior.Color = colorYesNoRed;
                    conditionNo.Font.Color = colorYesNoRedFont;

                    conditionYes.Interior.Color = colorYesNoGreen;
                    conditionYes.Font.Color = colorYesNoGreenFont;
            conditionYes.Font.Bold = true;
            conditionNo.Font.Bold = true;
            columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        private static void FormatR(Excel.Range columnRange)
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
            Excel.FormatCondition conditionR1 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ColorConverter cc = new ColorConverter();
            conditionR1.Interior.Color = c1;
            conditionR1.Font.Color = colorWhite;

            // R2
            Excel.FormatCondition conditionR2 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR2.Interior.Color = c2;
            conditionR2.Font.Color = colorWhite;

            // R3
            Excel.FormatCondition conditionR3 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R3", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR3.Interior.Color = c3;
            conditionR3.Font.Color = colorWhite;

            // R4
            Excel.FormatCondition conditionR4 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R4", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR4.Interior.Color = c4;
            conditionR4.Font.Color = colorWhite;

            // R5
            Excel.FormatCondition conditionR5 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R5", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR5.Interior.Color = c5;
            conditionR5.Font.Color = colorWhite;

            // R6
            Excel.FormatCondition conditionR6 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R6", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR6.Interior.Color = c6;
            conditionR6.Font.Color = colorWhite;

            // R7
            Excel.FormatCondition conditionR7 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R7", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR7.Interior.Color = c7;
            conditionR7.Font.Color = colorWhite;

            // R8
            Excel.FormatCondition conditionR8 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R8", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR8.Interior.Color = c8;
            conditionR8.Font.Color = colorWhite;

            // R9
            Excel.FormatCondition conditionR9 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R9", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR9.Interior.Color = c9;
            conditionR9.Font.Color = colorWhite;

            // R10
            Excel.FormatCondition conditionR10 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R10", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR10.Interior.Color = c10;
            conditionR10.Font.Color = colorWhite;

            // R11
            Excel.FormatCondition conditionR11 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R11", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR11.Interior.Color = c11;
            conditionR11.Font.Color = colorWhite;

            // R12
            Excel.FormatCondition conditionR12 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R12", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR12.Interior.Color = c12;
            conditionR12.Font.Color = colorWhite;

            // R13
            Excel.FormatCondition conditionR13 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R13", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR13.Interior.Color = c13;
            conditionR13.Font.Color = colorWhite;

            // R14
            Excel.FormatCondition conditionR14 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R14", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR14.Interior.Color = c14;
            conditionR14.Font.Color = colorWhite;

            // R15
            Excel.FormatCondition conditionR15 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R15", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR15.Interior.Color = c15;
            conditionR15.Font.Color = colorWhite;

            // R98
            Excel.FormatCondition conditionR98 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R98", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR98.Interior.Color = colorBlack;
            conditionR98.Font.Color = colorWhite;

            // R99
            Excel.FormatCondition conditionR99 = (Excel.FormatCondition)columnRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                   Excel.XlFormatConditionOperator.xlEqual, "R99", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            conditionR99.Interior.Color = colorDarkGrey;
            conditionR99.Font.Color = colorWhite;

            columnRange.Font.Bold = true;
            columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        public static void ExecuteToggleProperties(Excel.Application app)
        {
            try
            {
                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    int headerRow = 0;
                    int propertiesRow = 0;

                    Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    headerRow = headerRowRange.Row;
                    propertiesRow = headerRowRange.Row - 1;

                    if (propertiesRow > 0)
                    {
                        Excel.Worksheet ws = app.ActiveSheet;
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
