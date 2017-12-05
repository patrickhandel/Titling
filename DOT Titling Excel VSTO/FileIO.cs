﻿using System;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace DOT_Titling_Excel_VSTO
{
    class FileIO
    {
        public static void CreateRoadMapPDF(Worksheet ws)
        {
            string roadMapFileName = GetNewRoadMapFileName("pdf", ws.Name);
            const int xlQualityStandard = 0;
            ws.ExportAsFixedFormat(
                XlFixedFormatType.xlTypePDF,
                roadMapFileName,
                xlQualityStandard,
                true,
                false,
                Type.Missing,
                Type.Missing,
                true,
                Type.Missing);
        }

        public static void CreateRoadMapImage(Worksheet ws)
        {
            ws.Select();
            string fileName = GetNewRoadMapFileName("jpeg", ws.Name);
            string startRange = "A1";
            Range endRange = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = ws.get_Range(startRange, endRange);
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
    }
}
