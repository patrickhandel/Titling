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
            MergeDataIntoNewFile();
        }
        public static void MergeDataIntoNewFile()
        {
            string summary = "This is the summary";
            string JiraId = "DOTTITLNG-73";
            string id = JiraId.Replace("DOTTITLNG-", string.Empty);

            string template = @"C:\\Users\\patrick.handel\\Desktop\\MailMergeOut\\MyDoc.docx";
            string newfile = @"C:\\Users\\patrick.handel\\Desktop\\Exported\\" + summary + " (" + id + ").docx";
            File.Copy(template, newfile, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(newfile, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                docText = MergeField(docText, "Summary", summary);
                docText = MergeField(docText, "DOTTITLNG", JiraId);
                docText = MergeField(docText, "Epic", "Reset/Rollback - R2");
                docText = MergeField(docText, "Release", "Release");
                docText = MergeField(docText, "Sprint", "Sprint");
                docText = MergeField(docText, "Story1", "Story1");
                docText = MergeField(docText, "Story2", "Story2");
                docText = MergeField(docText, "Story3", "Story3");
                docText = MergeField(docText, "Description", "Description");
                docText = MergeField(docText, "Web Services", "Web Services");
                docText = MergeField(docText, "Date Approved", "12/15/2010");
                docText = MergeField(docText, "Document Date", DateTime.Now.ToShortDateString());

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private static String MergeField(string docText, string field, string newText)
        {
            Regex regexText = new Regex("{" + field + "}");
            return regexText.Replace(docText, newText); ;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
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
    }
}
