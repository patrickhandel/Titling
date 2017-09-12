using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DOT_Titling_Excel_VSTO
{
    class MailMerge
    {
        public static void ExecuteMailMerge()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;

                if (activeCell != null && activeWorksheet.Name == "Tickets")
                {
                    SSUtils.DoStandardStuff(app);
                    List<MailMergeFields> mailMergeFields = WorksheetPropertiesManager.GetMailMergeFields();
                    CreateMailMergeDocuments(app, activeWorksheet, selection, mailMergeFields);
                    SSUtils.DoStandardStuff(app);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void CreateMailMergeDocuments(Excel.Application app, Excel.Worksheet activeWorksheet, Excel.Range selection, List<MailMergeFields> mailMergeFields)
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

                        int jiraIDCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Story ID");
                        string jiraId = SSUtils.GetCellValue(activeWorksheet, row, jiraIDCol);
                        if (jiraId.Length > 10 && jiraId.Substring(0, 10) == "DOTTITLNG-")
                        {
                            Import.ExecuteImportSingleJiraTicket(jiraId);
                            string summary = string.Empty;
                            string epicID = string.Empty;
                            foreach (var mailMergeField in mailMergeFields)
                            {
                                string name = mailMergeField.Name;
                                string text = mailMergeField.Text;
                                int col = SSUtils.GetColumnFromHeader(activeWorksheet, mailMergeField.Text);
                                string value = SSUtils.GetCellValue(activeWorksheet, row, col);
                                if (name == "summary")
                                    summary = value;
                                if (name == "epicID")
                                    epicID = value;
                                foreach (Word.Field field in wordDocument.Fields)
                                {
                                    string fieldText = field.Code.Text;
                                    fieldText = fieldText.Replace("MERGEFIELD", String.Empty);
                                    fieldText = fieldText.Replace("\\* MERGEFORMAT", String.Empty);
                                    fieldText = fieldText.Replace(" ", String.Empty);

                                    if (fieldText == mailMergeField.Name)
                                    {
                                        field.Select();
                                        wordApp.Selection.TypeText(value);
                                    }
                                }
                            }

                            wordApp.Visible = false;
                            string newfile = SSUtils.GetNewFileName(summary, epicID);

                            wordDocument.TrackRevisions = true;
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
    }
}
