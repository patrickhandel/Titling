using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Jira = Atlassian.Jira;
using System.Windows.Forms;

namespace DOT_Titling_Excel_VSTO
{
    class MailMerge
    {
        public static void ExecuteMailMerge_DOT(Jira.Jira jira, Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                if (activeCell != null && ws.Name == "Issues")
                {
                    Excel.Range selection = app.Selection;
                    var mailMergeFields = WorksheetPropertiesManager.GetMailMergeFields();
                    CreateMailMergeDocuments(jira, app, ws, selection, mailMergeFields);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void CreateMailMergeDocuments(Jira.Jira jira, Excel.Application app, Excel.Worksheet ws, Excel.Range selection, List<MailMergeFields> mailMergeFields)
        {
            try
            {
                object template = @ThisAddIn.InputDir + "\\Requirement.docx";
                Word.Application wordApp = new Word.Application();
                Word.Document wordDocument = new Word.Document();
                wordApp.Visible = false;

                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (ws.Rows[row].EntireRow.Height != 0)
                    {
                        wordDocument = wordApp.Documents.Add(Template: template);

                        int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                        string issueID = SSUtils.GetCellValue(ws, row, issueIDCol);
                        string projectKey = ThisAddIn.ProjectKeyDOT;
                        if (issueID.Length > 10 && issueID.Substring(0, 10) == projectKey + "-")
                        {
                            JiraShared.ExecuteUpdateRowBeforeOperation(jira, app, ws, issueID, "Issue ID");
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
                                if (MessageBox.Show("Open " + newfile + "?", issueID, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
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
