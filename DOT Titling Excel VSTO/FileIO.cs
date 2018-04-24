using System;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class FileIO
    {
        public static void CreateRoadMapPDF(Excel.Worksheet ws)
        {
            string roadMapFileName = GetNewRoadMapFileName("pdf", ws.Name);
            const int xlQualityStandard = 0;
            ws.ExportAsFixedFormat(
                Excel.XlFixedFormatType.xlTypePDF,
                roadMapFileName,
                xlQualityStandard,
                true,
                false,
                Type.Missing,
                Type.Missing,
                true,
                Type.Missing);
        }

        public static void CreateRoadMapImage(Excel.Worksheet ws)
        {
            ws.Select();
            string fileName = GetNewRoadMapFileName("jpeg", ws.Name);
            string startRange = "A1";
            Excel.Range endRange = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = ws.get_Range(startRange, endRange);
            range.Copy();
            System.Drawing.Image imgRange1 = Clipboard.GetImage();
            imgRange1.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        public static string GetNewMailMergeFileName(string summary, string epicID)
        {
            return @ThisAddIn.OutputDir + "\\" + "Epic ID " + epicID.Trim() + " " + GetValidFileName(summary.Trim() + ".docx");
        }

        public static string GetNewRoadMapFileName(string ext, string wsName)
        {
            string dt = DateTime.Now.ToString("yyyy-MM-dd");
            return @ThisAddIn.RoadMapDir + "\\" + dt + " " + wsName + "." + ext;
        }

        public static string GetValidFileName(string text)
        {
            text = text.Replace('\'', ' '); // U+2019 right single quotation mark
            text = text.Replace('"', ' '); // U+201D right double quotation mark
            text = text.Replace('/', ' ');  // U+2044 fraction slash
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                text = text.Replace(c, ' ');
            }
            return text;
        }

        public static string GetHomePath()
        {
            // Not in .NET 2.0
            // System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (System.Environment.OSVersion.Platform == System.PlatformID.Unix)
                return System.Environment.GetEnvironmentVariable("HOME");

            return System.Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
        }

        public static string GetDownloadFolderPath()
        {
            if (Environment.OSVersion.Platform == PlatformID.Unix)
            {
                string pathDownload = Path.Combine(GetHomePath(), "Downloads");
                return pathDownload;
            }

            return Convert.ToString(
                Microsoft.Win32.Registry.GetValue(
                     @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                    , "{374DE290-123F-4565-9164-39C4925E467B}"
                    , String.Empty
                )
            );
        }

    }
}
