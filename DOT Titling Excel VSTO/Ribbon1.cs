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
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;

            if (activeCell != null)
            {
                int row = activeCell.Row;
                int jiraIDCol = 6;



                string JiraId = CellGetStringValue(activeWorksheet, row, jiraIDCol);
                if (activeWorksheet.Name == "Stories" && JiraId.Substring(1, 11) == "DOTTITLING-")
                {
                    int epicCol = 1;
                    int summaryCol = 5;
                    int releaseCol = 11;
                    int sprintCol = 13;
                    int story1Col = 30;
                    int story2Col = 31;
                    int story3Col = 32;
                    int descriptionCol = 29;
                    int webServicesCol = 33;
                    int dateSubmittedCol = 27;
                    int dateApprovedCol = 28;

                    string summary = CellGetStringValue(activeWorksheet, row, summaryCol);
                    string epic = CellGetStringValue(activeWorksheet, row, epicCol);
                    string release = CellGetStringValue(activeWorksheet, row, releaseCol);
                    string sprint = CellGetStringValue(activeWorksheet, row, sprintCol);
                    string story1 = CellGetStringValue(activeWorksheet, row, story1Col);
                    string story2 = CellGetStringValue(activeWorksheet, row, story2Col);
                    string story3 = CellGetStringValue(activeWorksheet, row, story3Col);
                    string description = CellGetStringValue(activeWorksheet, row, descriptionCol);
                    string webServices = CellGetStringValue(activeWorksheet, row, webServicesCol);
                    string dateSubmited = CellGetStringValue(activeWorksheet, row, dateSubmittedCol);
                    string dateApproved = CellGetStringValue(activeWorksheet, row, dateApprovedCol);

                    string id = JiraId.Replace("DOTTITLNG-", string.Empty);
                    string template = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MailMergeOut\\MyDoc.docx";
                    string newfile = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Exported\\" + summary + " (" + id + ").docx";
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
                        docText = MergeField(docText, "Epic", epic);
                        docText = MergeField(docText, "Release", release);
                        docText = MergeField(docText, "Sprint", sprint);
                        docText = MergeField(docText, "Story1", story1);
                        docText = MergeField(docText, "Story2", story2);
                        docText = MergeField(docText, "Story3", story3);
                        docText = MergeField(docText, "Description", description);
                        docText = MergeField(docText, "Web Services", webServices);
                        docText = MergeField(docText, "Date Submitted", dateSubmited);
                        docText = MergeField(docText, "Date Approved", dateApproved);
                        docText = MergeField(docText, "Document Date", DateTime.Now.ToShortDateString());

                        using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                        {
                            sw.Write(docText);
                        }
                    }
                }
            }
        }

        public static string CellGetStringValue(Excel.Worksheet sheet, int row, int column)
        {
            var result = string.Empty;

            if (sheet != null)
            {
                var rng = sheet.Cells[row, column] as Excel.Range;

                if (rng != null)
                    result = (string)rng.Text;
            }

            return result;
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
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;

            //Filtered selection
            for (int rowIndex = selection.Row; rowIndex < selection.Row + selection.Rows.Count; rowIndex++)
            {
                if (activeWorksheet.Rows[rowIndex].EntireRow.Height != 0)
                {
                    string sval = activeWorksheet.Rows[rowIndex].Text;
                }
            }

            if (activeCell != null)
            {
                string sValue = activeCell.Value2.ToString();
                string sText = activeCell.Text;
            }
        }
    }
}
