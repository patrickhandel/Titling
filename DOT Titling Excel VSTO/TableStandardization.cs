﻿using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class TableStandardization
    {
        public static void ExecuteCleanupTable(Excel.Application app)
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

                    string colType;
                    string columnHeader;

                    // Format each cell in the table header row
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
                                case "TextLong":
                                    cell.EntireColumn.ColumnWidth = 40;
                                    cell.IndentLevel = 1;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "TextMedium":
                                    cell.EntireColumn.ColumnWidth = 20;
                                    cell.IndentLevel = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "TextShort":
                                    cell.EntireColumn.ColumnWidth = 15;
                                    cell.IndentLevel = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "TextTiny":
                                    cell.EntireColumn.ColumnWidth = 9;
                                    cell.IndentLevel = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "Number":
                                    cell.IndentLevel = 0;
                                    cell.EntireColumn.ColumnWidth = 9;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "Percent":
                                    cell.IndentLevel = 0;
                                    cell.EntireColumn.ColumnWidth = 9;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "Date":
                                    cell.IndentLevel = 0;
                                    cell.EntireColumn.ColumnWidth = 9;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                case "Error":
                                case "YesNoGreen":
                                case "YesNoRed":
                                case "YesNo":
                                    cell.EntireColumn.ColumnWidth = 9;
                                    cell.IndentLevel = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    columnRange.Font.Bold = true;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;

                                    //columnRange.FormatConditions.Delete();
                                    //FormatCondition condition =
                                    //       (FormatCondition)columnRange.FormatConditions.Add(XlFormatConditionType.xlCellValue,
                                    //       XlFormatConditionOperator.xlEqual, "x",
                                    //       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    //condition.Interior.PatternColorIndex = Constants.xlAutomatic;
                                    //condition.Interior.TintAndShade = 0;

                                    //if (colType == "YesNoGreen")
                                    //{
                                    //    condition.Interior.Color = XlRgbColor.rgbLightGreen;
                                    //    condition.Font.Color = XlRgbColor.rgbDarkGreen;
                                    //}

                                    //if (colType == "YesNo" || colType == "YesNoRed")
                                    //{
                                    //    condition.Interior.Color = XlRgbColor.rgbLightPink;
                                    //    condition.Font.Color = XlRgbColor.rgbWhite;
                                    //}

                                    //condition.StopIfTrue = false;
                                    break;
                                case "Hidden":
                                    cell.IndentLevel = 0;
                                    cell.EntireColumn.ColumnWidth = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
                                default:
                                    cell.EntireColumn.ColumnWidth = 15;
                                    cell.IndentLevel = 0;
                                    columnRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    columnRange.Font.Bold = false;
                                    columnRange.Font.Italic = false;
                                    columnRange.Font.Underline = false;
                                    break;
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
