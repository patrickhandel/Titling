using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Windows.Forms;

namespace DOT_Titling_Excel_VSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = app.ActiveSheet;
            Excel.Range activeCell = app.ActiveCell;
            Excel.Range selection = app.Selection;

            if (activeCell != null && activeWorksheet.Name == "Stories")
            {
                Excel.Worksheet mmWorksheet = app.Worksheets.Add();
                ApplyMailMergeHeader(mmWorksheet);
                PopulateMailMergeWorksheet(activeWorksheet, selection, mmWorksheet);
                CopyMailMergeWorksheetToNewWorkbook(mmWorksheet, app);
            }
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MergeDataIntoNewFile();
        }


        public static void CopyMailMergeWorksheetToNewWorkbook(Excel.Worksheet ws, Excel.Application app)
        {
            string newFile = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\TempDoc" + DateTime.Now.ToFileTime() + ".xlsx";
            Excel.Workbook newWookbook = Globals.ThisAddIn.Application.Workbooks.Add(Type.Missing);

            int i = newWookbook.Worksheets.Count + 1;
            ws.Copy(Type.Missing, newWookbook.Worksheets[1]);
            app.DisplayAlerts = false;
            newWookbook.SaveAs(newFile,  Excel.XlSaveAsAccessMode.xlNoChange);
            ws.Delete();
            app.DisplayAlerts = true;
        }

        public static void ApplyMailMergeHeader(Excel.Worksheet ws)
        {
            string[] sFields;
            string sHeader = "jiraID, summary, epic, release, sprint, story1, story2, story3, description, webServices, dateSubmited, dateApproved";
            sFields = sHeader.Split(',');
            for (int i = 0; i < sFields.Length; i++)
            {
                ws.Cells[1, i + 1] = sFields[i].ToString();
            }
        }

        public static void PopulateMailMergeWorksheet(Excel.Worksheet storiesWorksheet, Excel.Range selection, Excel.Worksheet mmWorksheet)
        {
            //Populate Data
            int  mergeRow = 2;
            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (storiesWorksheet.Rows[row].EntireRow.Height != 0)
                {
                    //string sval = activeWorksheet.Rows[row].Text;
                    int jiraIDCol = 6;
                    string jiraID = CellGetStringValue(storiesWorksheet, row, jiraIDCol);
                    if (jiraID.Substring(0, 10) == "DOTTITLNG-")
                    {
                        int epicCol = 1;
                        int summaryCol = 5;
                        int releaseCol = 11;
                        int sprintCol = 13;
                        int dateApprovedCol = 28;
                        int dateSubmittedCol = 29;
                        int descriptionCol = 30;
                        int story1Col = 31;
                        int story2Col = 32;
                        int story3Col = 33;
                        int webServicesCol = 34;

                        string summary = CellGetStringValue(storiesWorksheet, row, summaryCol);
                        string epic = CellGetStringValue(storiesWorksheet, row, epicCol);
                        string release = CellGetStringValue(storiesWorksheet, row, releaseCol);
                        string sprint = CellGetStringValue(storiesWorksheet, row, sprintCol);
                        string story1 = CellGetStringValue(storiesWorksheet, row, story1Col);
                        string story2 = CellGetStringValue(storiesWorksheet, row, story2Col);
                        string story3 = CellGetStringValue(storiesWorksheet, row, story3Col);
                        string description = CellGetStringValue(storiesWorksheet, row, descriptionCol);
                        string webServices = CellGetStringValue(storiesWorksheet, row, webServicesCol);
                        string dateSubmited = CellGetStringValue(storiesWorksheet, row, dateSubmittedCol);
                        string dateApproved = CellGetStringValue(storiesWorksheet, row, dateApprovedCol);

                        mmWorksheet.Cells[mergeRow, 1] = jiraID;
                        mmWorksheet.Cells[mergeRow, 2] = summary;
                        mmWorksheet.Cells[mergeRow, 3] = epic;
                        mmWorksheet.Cells[mergeRow, 4] = release;
                        mmWorksheet.Cells[mergeRow, 5] = sprint;
                        mmWorksheet.Cells[mergeRow, 6] = story1;
                        mmWorksheet.Cells[mergeRow, 7] = story2;
                        mmWorksheet.Cells[mergeRow, 8] = story3;
                        mmWorksheet.Cells[mergeRow, 9] = description;
                        mmWorksheet.Cells[mergeRow, 10] = webServices;
                        mmWorksheet.Cells[mergeRow, 11] = dateSubmited;
                        mmWorksheet.Cells[mergeRow, 12] = dateApproved;

                        mergeRow++;
                    }
                }
                mmWorksheet.Rows.RowHeight = 15;
                mmWorksheet.Columns.AutoFit();
            }
        }

        private void CreateMailMergeExcelDataFile(object oDataFile)
        {
            try
            {
                Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Worksheet mmWorksheet = Globals.ThisAddIn.Application.Sheets["Mail Merge"];
                Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
                Excel.Range selection = Globals.ThisAddIn.Application.Selection;

                //jiraID
                //summary
                //epic
                //release
                //sprint
                //story1
                //story2
                //story3
                //description
                //webServices
                //dateSubmited
                //dateApproved

                string[] sFields, sRecord;
                string sHeader = "jiraID, summary, epic, release, sprint, story1, story2, story3, description, webServices, dateSubmited, dateApproved";
                string sFirstRecord = "John,Roy,31 New street,320009";
                object oQuery = "SELECT * FROM `Sheet1$`";

                sFields = sHeader.Split(',');
                sRecord = sFirstRecord.Split(',');
                ////writing in excel you Can use datatable and Get the records and loop.here for sample i have writing keeping two strings 
                for (int i = 0; i < sFields.Length; i++)
                {
                    activeWorksheet.Cells[1, i + 1] = sFields[i].ToString();
                }
                for (int j = 0; j < sRecord.Length; j++)
                {
                    activeWorksheet.Cells[2, j + 1] = sRecord[j].ToString();
                }

                //saving the excel workbook
                //excelwrbook.SaveAs(oName, MSExcel.XlFileFormat.xlTemplate, objMissing, objMissing, objMissing, objMissing, MSExcel.XlSaveAsAccessMode.xlExclusive, objMissing, objMissing, objMissing, objMissing, objMissing);
                //excelapp.Quit();

                ////opening the excel to act as a datasource for word mail merge
                //wrdDoc.MailMerge.OpenDataSource("C:\\Users\\patri\\Dropbox\\Desktop\\TempDoc.xls", ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                //    ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref oQuery,
                //    ref objMissing, ref objMissing, ref objMissing);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void PerformMailMerge()
        { 
            try
            {
                try
                {
                    Object oMissing = System.Reflection.Missing.Value;
                    Object oFalse = false;
                    Object oDate = "dddd, MMMM dd, yyyy";
                    Object oDataFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\TempDoc.xls";

                    Word.Application wordApp = new Word.Application();
                    Word.Document wordDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    Word.Selection wordSelection;
                    Word.MailMerge wordMailMerge;
                    Word.MailMergeFields wordMailMergeFields;

                    // Create an instance of Word  and make it visible.
                    wordApp.Visible = true;
                    wordDoc.Activate();
                    wordDoc.Select();
                    wordSelection = wordApp.Selection;
                    wordMailMerge = wordDoc.MailMerge;

                    // Create a paragraph
                    wordApp.Selection.TypeText("This is some text in my new Word document.");
                    wordApp.Selection.TypeParagraph();

                    // Create a MailMerge Data file using excel            
                    //CreateMailMergeExcelDataFile(oDataFile);

                    // Create a string and insert it into the document.
                    wordSelection.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSelection.TypeText("Mail Merge");

                    //Add Two Lines
                    wordApp.Selection.TypeParagraph();
                    wordApp.Selection.TypeParagraph();

                    // Insert merge data.
                    wordSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordMailMergeFields = wordMailMerge.Fields;
                    wordMailMergeFields.Add(wordSelection.Range, "jiraID");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "summary");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "epic");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "release");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "sprint");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "story1");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "story2");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "story3");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "webServices");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "dateSubmited");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "dateApproved");
                    wordSelection.TypeParagraph();
                    wordMailMergeFields.Add(wordSelection.Range, "description");

                    // Perform mail merge.
                    wordMailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
                    wordMailMerge.Execute(ref oFalse);

                    // Close the original form document.
                    wordDoc.Saved = true;
                    wordDoc.Close(ref oFalse, ref oMissing, ref oMissing);

                    // Makes the merged doc visible
                    wordApp.Visible = true;

                    // Release References.
                    wordSelection = null;
                    wordMailMerge = null;
                    wordMailMergeFields = null;
                    wordDoc = null;
                    wordApp = null;

                    // Clean up temp file.
                    //System.IO.File.Delete(oDataFile);

                    //KEEP
                    //wordSelection.TypeText(StrToAdd);
                    //wordSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //wordSelection.InsertDateTime(ref oDate, ref oFalse, ref oMissing, ref oMissing, ref oMissing);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error :" + ex);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void MergeDataIntoNewFile()
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;

            if (activeCell != null && activeWorksheet.Name == "Stories")
            {
                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (activeWorksheet.Rows[row].EntireRow.Height != 0)
                    {
                        //string sval = activeWorksheet.Rows[row].Text;
                        int jiraIDCol = 6;
                        string JiraId = CellGetStringValue(activeWorksheet, row, jiraIDCol);
                        if (JiraId.Substring(0, 10) == "DOTTITLNG-")
                        {
                            int epicCol = 1;
                            int summaryCol = 5;
                            int releaseCol = 11;
                            int sprintCol = 13;
                            int dateApprovedCol = 28;
                            int dateSubmittedCol = 29;
                            int descriptionCol = 30;
                            int story1Col = 31;
                            int story2Col = 32;
                            int story3Col = 33;
                            int webServicesCol = 34;

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

                            description = description.Replace("<", "[");
                            description = description.Replace("/>", "/]");
                            description = description.Replace("\r\n\r\n\r\n\r\n", "\r\n\r\n\r\n");
                            description = description.Replace("\r\n\r\n\r\n", "\r\n\r\n");
                            description = description.Replace("\r\n", "<w:cr/>");
                            description = description.Replace("â€¢\t", "-");
                            description = description.Replace("â€¢", "");

                            // Remove URLs
                            foreach (Match item in Regex.Matches(description, @"(http|ftp|https):\/\/([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?"))
                            {
                                description = description.Replace(item.Value, "{URL REMOVED}");
                            }

                            string id = JiraId.Replace("DOTTITLNG-", string.Empty);
                            string template = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MailMergeOut\\MyDoc.docx";
                            string newfile = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Exported\\" + SantizeFilename(summary + " (" + id + ").docx");

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

                                if (selection.Rows.Count == 1)
                                { 
                                    if (MessageBox.Show("Open " + newfile + "?", JiraId, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                                    {
                                        System.Diagnostics.Process.Start(newfile);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        static string SantizeFilename(string key)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (var c in key)
            {
                var invalidCharIndex = -1;
                for (var i = 0; i < invalidChars.Length; i++)
                {
                    if (c == invalidChars[i])
                    {
                        invalidCharIndex = i;
                    }
                }
                if (invalidCharIndex > -1)
                {
                    sb.Append("_").Append(invalidCharIndex);
                    continue;
                }

                if (c == '_')
                {
                    sb.Append("__");
                    continue;
                }

                sb.Append(c);
            }
            return sb.ToString();
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
