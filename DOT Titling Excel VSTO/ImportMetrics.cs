using Microsoft.Win32;
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
        public static bool Import(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                if ((ws.Name == "Metrics"))
                {
                    int headerRow = PrepTable(app, ws);

                    string f = GetTextFileName();
                    int counter = 0;
                    string line;
                    int row = 0;
                    // Read the file and display it line by line.  
                    System.IO.StreamReader file = new System.IO.StreamReader(f);
                    while ((line = file.ReadLine()) != null)
                    {
                        if (line.Left(5) == "Time:")
                        {
                            row++;
                            SetDateTime(ws, headerRow, line, row);
                        }
                        if (line.Left(6) == "Weight")
                        {
                            SetWeight(ws, headerRow, line, row);
                        }
                        if (line.Left(10) == "Body Water")
                        {
                            SetBodyWater(ws, headerRow, line, row);
                        }
                        if (line.Left(8) == "Body Fat")
                        {
                            SetBodyFat(ws, headerRow, line, row);
                        }
                        //Bone Mass
                        if (line.Left(9) == "Bone Mass")
                        {
                            SetBoneMass(ws, headerRow, line, row);
                        }
                        //BMI
                        if (line.Left(3) == "BMI")
                        {
                            SetBMI(ws, headerRow, line, row);
                        }
                        //Visceral Fat
                        if (line.Left(12) == "Visceral Fat")
                        {
                            SetVisceralFat(ws, headerRow, line, row);
                        }
                        //BMR
                        if (line.Left(3) == "BMR")
                        {
                            SetBMR(ws, headerRow, line, row);
                        }
                        //Muscle Mass
                        if (line.Left(11) == "Muscle Mass")
                        {
                            SetMuscleMass(ws, headerRow, line, row);
                        }
                        counter++;
                    }
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

        private static void SetDateTime(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string dateTime = GetDateTime(line);
            int col = SSUtils.GetColumnFromHeader(ws, "DateTime");
            if (row != 1)
            {
                Excel.Range rToInsert = ws.get_Range(String.Format("{0}:{0}", headerRow + row), Type.Missing);
                rToInsert.Insert();
            }
            SSUtils.SetCellValue(ws, headerRow + row, col, dateTime);
        }

        private static void SetWeight(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string weight = GetWeight(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Weight");
            SSUtils.SetCellValue(ws, headerRow + row, col, weight);
        }

        private static void SetBodyWater(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetBodyWater(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Body Water");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetBodyFat(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetBodyFat(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Body Fat");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetBoneMass(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetBoneMass(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Bone Mass");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetVisceralFat(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetVisceralFat(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Visceral Fat");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetMuscleMass(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetMuscleMass(line);
            int col = SSUtils.GetColumnFromHeader(ws, "Muscle Mass");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetBMR(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetBMR(line);
            int col = SSUtils.GetColumnFromHeader(ws, "BMR");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
        }

        private static void SetBMI(Excel.Worksheet ws, int headerRow, string line, int row)
        {
            string value = GetBMI(line);
            int col = SSUtils.GetColumnFromHeader(ws, "BMI");
            SSUtils.SetCellValue(ws, headerRow + row, col, value);
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

        private static string GetDateTime(string line)
        {
            line = line.Replace("Time:", "");
            int comma1 = line.IndexOf(',');
            int comma2 = line.IndexOf(',', line.IndexOf(',') + 1);
            string toRemove = line.Substring(comma1, comma2 - comma1 + 1);
            line = line.Replace(toRemove, "");
            return line;
        }

        private static string GetWeight(string line)
        {
            line = line.Replace("Weight  ", "");
            int index = line.IndexOf("lb ");
            line = line.Left(index);
            return line;
        }

        private static string GetBodyWater(string line)
        {
            line = line.Replace("Body Water  ", "");
            int index = line.IndexOf("%");
            line = line.Left(index);
            decimal percent = Decimal.Parse(line) * (decimal).01;
            return percent.ToString();
        }

        private static string GetBodyFat(string line)
        {
            line = line.Replace("Body Fat  ", "");
            int index = line.IndexOf("%");
            line = line.Left(index);
            decimal percent = Decimal.Parse(line) * (decimal).01;
            return percent.ToString();
        }

        private static string GetBoneMass(string line)
        {
            line = line.Replace("Bone Mass  ", "");
            int index = line.IndexOf("lb");
            line = line.Left(index);
            return line;
        }

        private static string GetBMI(string line)
        {
            line = line.Replace("BMI  ", "");
            int index = line.IndexOf(" ");
            line = line.Left(index);
            return line;
        }

        private static string GetVisceralFat(string line)
        {
            line = line.Replace("Visceral Fat  ", "");
            int index = line.IndexOf(" ");
            line = line.Left(index);
            return line;
        }

        private static string GetBMR(string line)
        {
            line = line.Replace("BMR  ", "");
            int index = line.IndexOf("Kcal");
            line = line.Left(index);
            return line;
        }

        private static string GetMuscleMass(string line)
        {
            line = line.Replace("Muscle Mass  ", "");
            int index = line.IndexOf("lb");
            line = line.Left(index);
            return line;
        }

        private static string GetTextFileName()
        {
            string downloadPath = FileIO.GetDownloadFolderPath();
            return downloadPath + "\\" + "1byone wellness 2.0.txt";

            // Displays an OpenFileDialog so the user can select a Cursor.  
            //System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            //openFileDialog1.Filter = "Cursor Files|*.cur";
            //openFileDialog1.Title = "Select a Cursor File";

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .CUR file was selected, open it.  
            //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            // Assign the cursor in the Stream to the Form's Cursor property.  
            //this.Cursor = new Cursor(openFileDialog1.OpenFile());
            //}
        }
    }
}
