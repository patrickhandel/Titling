﻿using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace DOT_Titling_Excel_VSTO
{
    class History
    {
        public static void ExecuteUpdateDeveloper_DOT()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorksheet = app.ActiveSheet;
                var activeCell = app.ActiveCell;
                var selection = app.Selection;                
                if (activeCell != null && activeWorksheet.Name == "Sprint Results")
                {
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    GetDeveloperFromHistory(app, activeWorksheet, selection, listofProjects);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        public async static void GetDeveloperFromHistory(Excel.Application app, Worksheet activeWorksheet, Range selection, List<string> listofProjects)
        {
            try
            {
                // Get Dev Column (14)
                int devCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Dev"); 
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                {
                    if (activeWorksheet.Rows[row].EntireRow.Height != 0)
                    {
                        // Get Jira ID (1)
                        int jiraIDCol = SSUtils.GetColumnFromHeader(activeWorksheet, "Ticket ID");
                        string jiraId = SSUtils.GetCellValue(activeWorksheet, row, jiraIDCol);
                        if (jiraId.Length > 10 && jiraId.Substring(0, 10) == listofProjects[0] + "-")
                        {
                            var historyItems = await ThisAddIn.GlobalJira.Issues.GetChangeLogsAsync(jiraId);
                            var developers = WorksheetPropertiesManager.GetDevelopers();
                            string devs = string.Empty;
                            foreach (var item in historyItems)
                            {
                                string thisDev  = item.Author.ToString();
                                foreach (var dev in developers)
                                {
                                    if (dev.DevName == thisDev && !devs.Contains(dev.ReplaceWith))
                                        devs = devs + " " + dev.ReplaceWith;
                                }
                            }
                            SSUtils.SetCellValue(activeWorksheet, row, devCol, devs.Trim(), "?");
                        }
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