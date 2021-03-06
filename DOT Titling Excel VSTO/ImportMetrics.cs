﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class Metrics
    {
        //Public Methods

        //public static bool ExecuteCalculator()
        //{
        //    WebServices.CalculatorSoapClient client = new WebServices.CalculatorSoapClient("CalculatorSoap");
        //    int result = client.Add(5, 8);

        //}


        public static bool Import(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if ((ws.Name == "Metrics"))
                {
                    int headerRow = SSUtils.GetHeaderRow(ws);
                    string textFileName = GetTextFileName();
                    string line;
                    System.IO.StreamReader file = new System.IO.StreamReader(textFileName);
                    int row = 0;
                    while ((line = file.ReadLine()) != null)
                    {
                        if (line.Left(5) == "Time:")
                        {
                            row = SSUtils.GetFooterRow(ws);
                            Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0}", row), Type.Missing);
                            rToInsert.Insert();
                            SetFormulas(ws, row);
                            SetValue(ws, line, row, "Time:", ",");
                        }
                        if (line.Left(6) == "Weight")
                        {
                            SetValue(ws, line, row, "Weight", "lb");
                        }
                        if (line.Left(10) == "Body Water")
                        {
                            SetValue(ws, line, row, "Body Water", "%");
                        }
                        if (line.Left(8) == "Body Fat")
                        {
                            SetValue(ws, line, row, "Body Fat", "%");
                        }
                        if (line.Left(9) == "Bone Mass")
                        {
                            SetValue(ws, line, row, "Bone Mass", "lb");
                        }
                        if (line.Left(3) == "BMI")
                        {
                            SetValue(ws, line, row, "BMI", "");
                        }
                        if (line.Left(12) == "Visceral Fat")
                        {
                            SetValue(ws, line, row, "Visceral Fat", "");
                        }
                        if (line.Left(3) == "BMR")
                        {
                            SetValue(ws, line, row, "BMR", "Kcal");
                        }
                        if (line.Left(11) == "Muscle Mass")
                        {
                            SetValue(ws, line, row, "Muscle Mass", "lb");
                        }
                    }

                    //Remove duplicates based on the first column
                    Excel.Range rngToDedupe = ws.get_Range("MetricsData", Type.Missing);
                    object cols = new object[] { 1, 1 };
                    rngToDedupe.RemoveDuplicates(cols, Excel.XlYesNoGuess.xlYes);

                    file.Close();
                    bool success = true;
                    return success;
                }
                else
                {
                    MessageBox.Show(ws.Name + " can't be updated.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

        private static void SetFormulas(Excel.Worksheet ws, int row)
        {
            int col = SSUtils.GetColumnFromHeader(ws, "Target Weight 1");
            SSUtils.SetCellFormula(ws, row, col, "=TargetWeight1");
            col = SSUtils.GetColumnFromHeader(ws, "Target Weight 2");
            SSUtils.SetCellFormula(ws, row, col, "=TargetWeight2");
            col = SSUtils.GetColumnFromHeader(ws, "Target BMI");
            SSUtils.SetCellFormula(ws, row, col, "=TargetBMI1");
        }

        private static void SetValue(Excel.Worksheet ws, string line, int row, string column, string delimeter)
        {
            string value = string.Empty;
            if (column == "Time:")
            {
                value = GetDateTimeFromLine(line, column, delimeter);
                column = "DateTime";
            }
            else
            {
                value = GetValueFromLine(line, column, delimeter);
            }
            if (delimeter == "%")
            {
                decimal percent = Decimal.Parse(value) * (decimal).01;
                value = percent.ToString();
            }
            int col = SSUtils.GetColumnFromHeader(ws, column);
            SSUtils.SetCellValue(ws, row, col, value);
        }

        private static string GetDateTimeFromLine(string line, string value, string delimeter)
        {
            line = line.Replace(value, "");
            int delim1 = line.IndexOf(delimeter);
            int delim2 = line.IndexOf(delimeter, line.IndexOf(delimeter) + 1);
            string toRemove = line.Substring(delim1, delim2 - delim1 + 1);
            line = line.Replace(toRemove, "");
            return line;
        }

        private static string GetValueFromLine(string line, string value, string delimeter)
        {
            line = line.Replace(value + "  ", "");
            int index = line.IndexOf(delimeter + " ");
            line = line.Left(index);
            return line;
        }

        private static string GetTextFileName()
        {
            string downloadPath = FileIO.GetDownloadFolderPath();
            string textFile = FileIO.GetLastFileInDirectory(downloadPath, "1byone wellness*.txt");
            return downloadPath + "\\" + textFile;
        }

        private static int PrepTable(Excel.Application app, Excel.Worksheet ws)
        {
            // Get the header row
            string headerRangeName = SSUtils.GetSelectedTableHeader(app);
            Excel.Range headerRowRange = app.get_Range(headerRangeName, Type.Missing);
            int headerRow = headerRowRange.Row;

            // Get the footer row
            string footerRangeName = SSUtils.GetSelectedTableFooter(app);
            Excel.Range footerRowRange = app.get_Range(footerRangeName, Type.Missing);
            int footerRow = footerRowRange.Row;

            // Delete existing rows 
            if (footerRow > headerRow + 2)
            {
                Excel.Range rToDelete = ws.get_Range(String.Format("{0}:{1}", headerRow + 1, footerRow - 1), Type.Missing);
                rToDelete.Delete();
            }

            return headerRow;
        }
    }
}
