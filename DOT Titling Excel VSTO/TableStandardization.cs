using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class TableStandardization
    {
        public enum StandardizationType
        {
            Thorough = 1,
            Light = 2
        };

        public static void ExecuteCleanupTable(Excel.Application app, StandardizationType type)
        {
            try
            {
                string tableRangeName = SSUtils.GetSelectedTable(app);
                string headerRangeName = SSUtils.GetSelectedTableHeader(app);
                if (headerRangeName != string.Empty)
                {
                    int column = 0;
                    int headerRow = 0;
                    int footerRow = 0;
                    int footerRowOffset = 0;

                    Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
                    headerRow = headerRowRange.Row;

                    // Get the footer row and format it
                    if (type == StandardizationType.Thorough)
                    {
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
                            Range offsetRowRange = headerRowRange.Offset[-1, 0];
                            if (offsetRowRange != null)
                            {
                                headerRowRange.Copy(Type.Missing);
                                offsetRowRange.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                    }

                    string colType;
                    string columnHeader;

                    // Format each cell in the table header row
                    foreach (Range cell in headerRowRange.Cells)
                    {
                        column = cell.Column;
                        columnHeader = cell.Value;
                        colType = cell.Offset[-1, 0].Value;

                        switch (colType)
                        {
                            case "TextLong":
                                cell.EntireColumn.ColumnWidth = 40;
                                cell.IndentLevel = 1;
                                break;
                            case "TextMedium":
                                cell.EntireColumn.ColumnWidth = 20;
                                break;
                            case "TextShort":
                                cell.EntireColumn.ColumnWidth = 15;
                                break;
                            case "TextTiny":
                                cell.EntireColumn.ColumnWidth = 9;
                                break;
                            case "Number":
                                cell.EntireColumn.ColumnWidth = 9;
                                break;
                            case "Percent":
                                cell.EntireColumn.ColumnWidth = 9;
                                break;
                            case "Date":
                                cell.EntireColumn.ColumnWidth = 10;
                                break;
                            case "Error":
                            case "YesNoGreen":
                            case "YesNoRed":
                            case "YesNo":
                                cell.EntireColumn.ColumnWidth = 7;
                                break;
                            case "MidLong":
                                cell.EntireColumn.ColumnWidth = 13;
                                break;
                            case "Release":
                                cell.EntireColumn.ColumnWidth = 7;
                                break;
                            case "Hidden":
                                cell.EntireColumn.ColumnWidth = 0;
                                break;
                            default:
                                cell.EntireColumn.ColumnWidth = 15;
                                break;
                        }

                        if (type == StandardizationType.Thorough)
                        {
                            string columnNameRange = tableRangeName + '[' + columnHeader + ']';
                            Range columnRange = app.get_Range(columnNameRange, Type.Missing);
                            if (columnRange != null)
                            {
                                cell.IndentLevel = 0;
                                columnRange.Font.Bold = false;
                                columnRange.Font.Italic = false;
                                columnRange.Font.Underline = false;
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                                switch (colType)
                                {
                                    case "TextLong":
                                        cell.EntireColumn.ColumnWidth = 40;
                                        cell.IndentLevel = 1;
                                        break;
                                    case "TextMedium":
                                        cell.EntireColumn.ColumnWidth = 20;
                                        break;
                                    case "TextShort":
                                        cell.EntireColumn.ColumnWidth = 15;
                                        break;
                                    case "TextTiny":
                                        cell.EntireColumn.ColumnWidth = 9;
                                        break;
                                    case "Number":
                                        cell.EntireColumn.ColumnWidth = 9;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                        break;
                                    case "Percent":
                                        cell.EntireColumn.ColumnWidth = 9;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                        columnRange.NumberFormat = "0%";
                                        break;
                                    case "Date":
                                        cell.EntireColumn.ColumnWidth = 10;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                        columnRange.NumberFormat = "m/d/yyyy";
                                        break;
                                    case "Error":
                                    case "YesNoGreen":
                                    case "YesNoRed":
                                    case "YesNo":
                                        cell.EntireColumn.ColumnWidth = 7;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                        columnRange.Font.Bold = true;
                                        break;
                                    case "MidLong":
                                        cell.EntireColumn.ColumnWidth = 13;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                        columnRange.Font.Bold = true;
                                        break;
                                    case "Release":
                                        cell.EntireColumn.ColumnWidth = 7;
                                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                        columnRange.Font.Bold = true;
                                        break;
                                    case "Hidden":
                                        cell.EntireColumn.ColumnWidth = 0;
                                        break;
                                    default:
                                        cell.EntireColumn.ColumnWidth = 15;
                                        break;
                                }
                            }
                        }
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

                    // Format the table offset row
                    headerRowRange.Offset[-1, 0].Font.Size = 9;
                    headerRowRange.EntireRow.Offset[-1, 0].Hidden = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
