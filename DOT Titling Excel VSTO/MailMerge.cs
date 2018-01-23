using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DOT_Titling_Excel_VSTO
{
    class MailMerge
    {
        public static void ExecuteMailMerge(Excel.Application app)
        {
            try
            {
                Worksheet activeWorksheet = app.ActiveSheet;
                Range activeCell = app.ActiveCell;
                var selection = app.Selection;
                if (activeCell != null && activeWorksheet.Name == "Tickets")
                {
                    var mailMergeFields = WorksheetPropertiesManager.GetMailMergeFields();
                    CreateMailMergeDocuments(app, activeWorksheet, selection, mailMergeFields);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void CreateMailMergeDocuments(Excel.Application app, Worksheet ws, Range selection, List<MailMergeFields> mailMergeFields)
        {
            try
            {
                object template = @ThisAddIn.InputDir + "\\Requirement.docx";
                var wordApp = new Word.Application();
                var wordDocument = new Word.Document();
                wordApp.Visible = false;

                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (ws.Rows[row].EntireRow.Height != 0)
                    {
                        wordDocument = wordApp.Documents.Add(Template: template);

                        int jiraIDCol = SSUtils.GetColumnFromHeader(ws, "Ticket ID");
                        string jiraId = SSUtils.GetCellValue(ws, row, jiraIDCol);
                        if (jiraId.Length > 10 && jiraId.Substring(0, 10) == "DOTTITLNG-")
                        {
                            ImportFromJira.ExecuteUpateTicketBeforeMailMerge(jiraId);
                            string summary = string.Empty;
                            string epicID = string.Empty;
                            foreach (var mailMergeField in mailMergeFields)
                            {
                                string name = mailMergeField.Name;
                                string text = mailMergeField.Text;
                                int col = SSUtils.GetColumnFromHeader(ws, mailMergeField.Text);
                                string value = SSUtils.GetCellValue(ws, row, col);
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
                                        if (value == string.Empty)
                                            value = " ";
                                        wordApp.Selection.TypeText(value);
                                    }
                                }
                            }

                            foreach (Word.Field field in wordDocument.Fields)
                            {
                                string fieldText = field.Code.Text;
                                fieldText = fieldText.Replace("MERGEFIELD", String.Empty);
                                fieldText = fieldText.Replace("\\* MERGEFORMAT", String.Empty);
                                fieldText = fieldText.Replace(" ", String.Empty);
                                if (fieldText == "dateDocumentCreated")
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(DateTime.Now.ToString("M/d/yyyy"));
                                }
                            }

                            wordApp.Visible = false;
                            string newfile = FileIO.GetNewMailMergeFileName(summary, epicID);

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
