using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DOT_Titling_Excel_VSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection;

            if (activeCell != null)
            {
                string sValue = activeCell.Value2.ToString();
                string sText = activeCell.Text;
                System.Windows.Forms.MessageBox.Show(sText);
            }
        }
        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                //{Summary}
                Regex regexText = new Regex("xxxSummaryxxx");
                docText = regexText.Replace(docText, "Reset by Envelope Number - After 8 PM");

                //{Epic}
                regexText = new Regex("{Epic}");
                docText = regexText.Replace(docText, "Reset/Rollback - R2");

                //{Story ID}
                regexText = new Regex("{Story ID}");
                docText = regexText.Replace(docText, "DOTTITLNG-73");

                //{Release}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Sprint}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Story1}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Story2}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Story3}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Description}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Web Services}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");

                //{Date Approved}
                regexText = new Regex("XXXXXXXXXX");
                docText = regexText.Replace(docText, "XXXXXXXXXXX");


                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}
