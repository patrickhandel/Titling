using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;

namespace DOT_Titling_Excel_VSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;

                if (activeCell != null && activeWorksheet.Name == "Stories")
                {
                    app.ScreenUpdating = false;

                    CreateWordDocuments(activeWorksheet, selection);

                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;

                if (activeCell != null && activeWorksheet.Name == "Stories")
                {
                    app.ScreenUpdating = false;
                    //Excel.Worksheet mmWorksheet = app.Worksheets.Add();
                    //MailMerge_CreateHeader(mmWorksheet);
                    //MailMerge_CreateData(activeWorksheet, selection, mmWorksheet);
                    //string dataFile = MailMerge_CreateDataFile(mmWorksheet, app);
                    //MailMerge_PerformMerge_Old(dataFile);
                    CreatedMergedDocuments(activeWorksheet, selection);
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void CreatedMergedDocuments(Excel.Worksheet activeWorksheet, Excel.Range selection)
        {
            try
            {
                Object oTemplate = @ThisAddIn.InputDir + "\\MyDocMerge.docx";
                var wordApp = new Word.Application();
                var wordDocument = new Word.Document();
                wordApp.Visible = false;

                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (activeWorksheet.Rows[row].EntireRow.Height != 0)
                    {
                        wordDocument = wordApp.Documents.Add(Template: oTemplate);

                        //string sval = activeWorksheet.Rows[row].Text;
                        int jiraIDCol = 6;
                        string jiraId = CellGetStringValue(activeWorksheet, row, jiraIDCol);
                        if (jiraId.Substring(0, 10) == "DOTTITLNG-")
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
                            string id = jiraId.Replace("DOTTITLNG-", string.Empty);

                            foreach (Microsoft.Office.Interop.Word.Field field in wordDocument.Fields)
                            {
                                if (field.Code.Text.Contains("jiraID"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(jiraId);
                                }
                                else if (field.Code.Text.Contains("summary"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(summary);
                                }
                                else if (field.Code.Text.Contains("epic"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(epic);
                                }
                                else if (field.Code.Text.Contains("release"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(release);
                                }
                                else if (field.Code.Text.Contains("sprint"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(sprint);
                                }
                                else if (field.Code.Text.Contains("story1"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(story1);
                                }
                                else if (field.Code.Text.Contains("story2"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(story2);
                                }
                                else if (field.Code.Text.Contains("story3"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(story3);
                                }
                                else if (field.Code.Text.Contains("description"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(description);
                                }
                                else if (field.Code.Text.Contains("webServices"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(webServices);
                                }
                                else if (field.Code.Text.Contains("dateSubmited"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(dateSubmited);
                                }
                                else if (field.Code.Text.Contains("dateApproved"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(dateApproved);
                                }
                            }
                            wordApp.Visible = false;

                            string newfile = @ThisAddIn.OutputDir + "\\" + MakeValidFilename(summary.Trim() + "( " + id.Trim() + ").docx");
                            wordDocument.SaveAs2(newfile);
                            wordDocument.Close(false);

                            if (selection.Rows.Count == 1)
                            {
                                if (MessageBox.Show("Open " + newfile + "?", jiraId, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                                {
                                    System.Diagnostics.Process.Start(newfile);
                                }
                            }
                        }
                    }
                }
                if (selection.Rows.Count > 1)
                {
                    if (MessageBox.Show("Open " + ThisAddIn.OutputDir + "?", selection.Rows.Count.ToString() + " Files Created", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(ThisAddIn.OutputDir);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void MailMerge_CreateHeader(Excel.Worksheet ws)
        {
            try
            {
                string[] sFields;
                string sHeader = "jiraID, summary, epic, release, sprint, story1, story2, story3, description, webServices, dateSubmited, dateApproved";
                sFields = sHeader.Split(',');
                for (int i = 0; i < sFields.Length; i++)
                {
                    ws.Cells[1, i + 1] = sFields[i].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void MailMerge_CreateData(Excel.Worksheet storiesWorksheet, Excel.Range selection, Excel.Worksheet mmWorksheet)
        {
            try
            {
                //Populate Data
                int mergeRow = 2;
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
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static string MailMerge_CreateDataFile(Excel.Worksheet ws, Excel.Application app)
        {
            try
            {
                string newFileName = ThisAddIn.OutputDir + "\\MailMergeData_" + DateTime.Now.ToFileTime() + ".xlsx";
                object newFile = newFileName;
                Excel.Workbook newWookbook = Globals.ThisAddIn.Application.Workbooks.Add(Type.Missing);
                ws.Copy(Type.Missing, newWookbook.Worksheets[1]);
                Excel.Worksheet newWorksheet = newWookbook.Worksheets[2];
                Excel.Worksheet toRemove = newWookbook.Worksheets[1];
                newWorksheet.Name = "MailMerge";
                app.DisplayAlerts = false;
                toRemove.Delete();
                newWookbook.SaveAs(newFile, Excel.XlFileFormat.xlOpenXMLWorkbook);
                newWookbook.Close();
                ws.Delete();
                app.DisplayAlerts = true;
                return newFileName;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return "";
            }
        }

        public static void MailMerge_PerformMerge_Old(string dataFile)
        { 
            try
            {
                try
                {
                    Object oMissing = System.Reflection.Missing.Value;
                    Object oFalse = false;
                    Object oDate = "dddd, MMMM dd, yyyy";
                    Object oDataFile = @dataFile;
                    Object oQuery = "SELECT * FROM `MailMerge$`";

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

                    //opening the excel to act as a datasource for word mail merge
                    wordDoc.MailMerge.OpenDataSource(dataFile, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oQuery,
                        ref oMissing, ref oMissing, ref oMissing);

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
                    System.IO.File.Delete(dataFile);
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

        public static void CreateWordDocuments(Excel.Worksheet activeWorksheet, Excel.Range selection)
        {
            try
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
                            string template = ThisAddIn.InputDir + "\\MyDoc.docx";
                            string newfile = ThisAddIn.OutputDir + "\\" + MakeValidFilename(summary.Trim() + " (" + id.Trim() + ").docx");

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
                if (selection.Rows.Count > 1)
                {
                    if (MessageBox.Show("Open " + ThisAddIn.OutputDir + "?", selection.Rows.Count.ToString() + " Files Created", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(ThisAddIn.OutputDir);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        static string MakeValidFilename(string text)
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

            return result + " ";
        }

        private static String MergeField(string docText, string field, string newText)
        {
            Regex regexText = new Regex("{" + field + "}");
            return regexText.Replace(docText, newText); ;
        }

    }
}
