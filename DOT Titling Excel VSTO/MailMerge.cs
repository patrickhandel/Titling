﻿using System;
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

                if (activeCell != null && activeWorksheet.Name == "Stories")
                {
                    app.ScreenUpdating = false;
                    CreateMailMergeDocuments(app, activeWorksheet, selection);
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public static void CreateMailMergeDocuments(Excel.Application app, Excel.Worksheet activeWorksheet, Excel.Range selection)
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

                        Dictionary<string, int> dict = new Dictionary<string, int>();
                        //dict = SetColumns(app);
                        dict = SetColumns1(app, activeWorksheet);

                        //string sval = activeWorksheet.Rows[row].Text;
                        int jiraIDCol = dict["jiraID"];
                        string jiraId = SSUtils.GetCellValue(activeWorksheet, row, jiraIDCol);
                        if (jiraId.Length > 10 && jiraId.Substring(0, 10) == "DOTTITLNG-")
                        {
                            int col_epic = dict["epic"];
                            int col_summary = dict["summary"];
                            int col_release = dict["release"];
                            int col_sprint = dict["sprint"];
                            int col_dateApproved = dict["dateApproved"];
                            int col_dateSubmitted = dict["dateSubmitted"];
                            int col_description = dict["description"];
                            int col_story1 = dict["story1"];
                            int col_story2 = dict["story2"];
                            int col_story3 = dict["story3"];
                            int col_webServices = dict["webServices"];
                            int col_epicID = dict["epicID"];
                            int col_storyCode = dict["storyCode"];

                            string summary = SSUtils.GetCellValue(activeWorksheet, row, col_summary);
                            string epic = SSUtils.GetCellValue(activeWorksheet, row, col_epic);
                            string release = SSUtils.GetCellValue(activeWorksheet, row, col_release);
                            string sprint = SSUtils.GetCellValue(activeWorksheet, row, col_sprint);
                            string story1 = SSUtils.GetCellValue(activeWorksheet, row, col_story1);
                            string story2 = SSUtils.GetCellValue(activeWorksheet, row, col_story2);
                            string story3 = SSUtils.GetCellValue(activeWorksheet, row, col_story3);
                            string description = SSUtils.GetCellValue(activeWorksheet, row, col_description);
                            string webServices = SSUtils.GetCellValue(activeWorksheet, row, col_webServices);
                            string epicID = SSUtils.GetCellValue(activeWorksheet, row, col_epicID).Trim();
                            string storyCode = SSUtils.GetCellValue(activeWorksheet, row, col_storyCode).Trim();
                            string dateSubmited = SSUtils.GetCellValue(activeWorksheet, row, col_dateSubmitted);
                            string dateApproved = SSUtils.GetCellValue(activeWorksheet, row, col_dateApproved);

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
                                else if (field.Code.Text.Contains("epicID"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(epicID);
                                }
                                else if (field.Code.Text.Contains("storyCode"))
                                {
                                    field.Select();
                                    wordApp.Selection.TypeText(storyCode);
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

        private static Dictionary<string, int> SetColumns1(Excel.Application app, Excel.Worksheet ws)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            string sHeaderRangeName = "StoryData[#Headers]";
            Excel.Range headerRowRange = (Excel.Range)ws.get_Range(sHeaderRangeName, Type.Missing);
            string header = "";
            int column = 1;

            foreach (Excel.Range cell in headerRowRange.Cells)
            {
                header = cell.Value;
                column = cell.Column;
                switch (header)
                {
                    case "Epic":
                        dict.Add("epic", column);
                        break;
                    case "Summary":
                        dict.Add("summary", column);
                        break;
                    case "Story ID":
                        dict.Add("jiraID", column);
                        break;
                    case "Story Release":
                        dict.Add("release", column);
                        break;
                    case "DOT Sprint":
                        dict.Add("sprint", column);
                        break;
                    case "Date Submitted to DOT":
                        dict.Add("dateSubmitted", column);
                        break;
                    case "Date Approved by DOT":
                        dict.Add("dateApproved", column);
                        break;
                    case "Description":
                        dict.Add("description", column);
                        break;
                    case "Story: As A":
                        dict.Add("story1", column);
                        break;
                    case "Story: I'd Like":
                        dict.Add("story2", column);
                        break;
                    case "Story: So That":
                        dict.Add("story3", column);
                        break;
                    case "Story Code":
                        dict.Add("storyCode", column);
                        break;
                    case "Epic ID":
                        dict.Add("epicID", column);
                        break;
                    case "DOT Web Services":
                        dict.Add("webServices", column);
                        break;
                    default:
                        break;
                }
            }
            return dict;
        }

        private static Dictionary<string, int> SetColumns(Excel.Application app)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            Excel.Workbook wb = app.ActiveWorkbook;

            var col_dateApproved = wb.Names.Item("col_dateApproved").RefersToRange.Value;
            dict.Add("dateApproved", (int)col_dateApproved);

            var col_dateSubmitted = wb.Names.Item("col_dateSubmitted").RefersToRange.Value;
            dict.Add("dateSubmitted", (int)col_dateSubmitted);

            var col_description = wb.Names.Item("col_description").RefersToRange.Value;
            dict.Add("description", (int)col_description);

            var col_epic = wb.Names.Item("col_epic").RefersToRange.Value;
            dict.Add("epic", (int)col_epic);

            var col_jiraID = wb.Names.Item("col_jiraID").RefersToRange.Value;
            dict.Add("jiraID", (int)col_jiraID);

            var col_release = wb.Names.Item("col_release").RefersToRange.Value;
            dict.Add("release", (int)col_release);

            var col_sprint = wb.Names.Item("col_sprint").RefersToRange.Value;
            dict.Add("sprint", (int)col_sprint);

            var col_story1 = wb.Names.Item("col_story1").RefersToRange.Value;
            dict.Add("story1", (int)col_story1);

            var col_story2 = wb.Names.Item("col_story2").RefersToRange.Value;
            dict.Add("story2", (int)col_story2);

            var col_story3 = wb.Names.Item("col_story3").RefersToRange.Value;
            dict.Add("story3", (int)col_story3);

            var col_summary = wb.Names.Item("col_summary").RefersToRange.Value;
            dict.Add("summary", (int)col_summary);

            var col_webServices = wb.Names.Item("col_webServices").RefersToRange.Value;
            dict.Add("webServices", (int)col_webServices);

            var col_epicID = wb.Names.Item("col_epicID").RefersToRange.Value;
            dict.Add("epicID", (int)col_epicID);

            var col_storyCode = wb.Names.Item("col_storyCode").RefersToRange.Value;
            dict.Add("storyCode", (int)col_storyCode);

            return dict;
        }
    }
}
