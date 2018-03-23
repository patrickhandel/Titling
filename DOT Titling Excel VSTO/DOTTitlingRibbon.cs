using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    public partial class DOTTitlingRibbon
    {
        // CLASS-SPECIFIC
        private void DOTTitlingRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Views_DOT_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private static Jira.Jira GetJira(Excel.Application app, string type)
        {
            try
            {
                if (type == "DOT") 
                {
                    return Jira.Jira.CreateRestClient(ThisAddIn.JiraSite, ThisAddIn.JiraUserName, ThisAddIn.JiraPassword);
                }
                else
                {
                    Excel.Worksheet wsUser = app.Sheets["User"];
                    Excel.Range rangeJiraUserName = wsUser.get_Range("JiraUserName");
                    Excel.Range rangeJiraPassword = wsUser.get_Range("JiraPassword");
                    string jiraUserName = rangeJiraUserName.Value2;
                    string jiraPassword = rangeJiraPassword.Value2;
                    return Jira.Jira.CreateRestClient(ThisAddIn.JiraSite, jiraUserName, jiraPassword);
                }
            }
            //catch (Exception ex)
            catch
            {
                MessageBox.Show("Error : Not a properly formatted workbook.");
                return null;
            }
        }

        // UPDATE, ADD, SAVE
        private void btnUpdateIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "DOT");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    JiraShared.ExecuteUpdateTable(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateIssues_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "Program");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = SSUtils.GetListOfProjects(app);
                    JiraShared.ExecuteUpdateTable(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAddIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "DOT");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    JiraShared.ExecuteAddNewRowsToTable(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAddIssues_Progam_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "Program");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = SSUtils.GetListOfProjects(app);
                    JiraShared.ExecuteAddNewRowsToTable(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateSelectedIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "DOT");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    JiraShared.ExecuteUpdateSelectedRows(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, "Selected Items Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateSelectedIssues_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "Program");

                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = SSUtils.GetListOfProjects(app);
                    JiraShared.ExecuteUpdateSelectedRows(jira, app, listofProjects);
                    SSUtils.EndExcelOperation(app, "Selected Items Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnSaveSelectedIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "DOT");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    bool multiple = JiraShared.ExecuteSaveSelectedCellsToJira(jira, app, listofProjects);
                    string msg = string.Empty;
                    if (multiple == true)
                    {
                        msg = "Selected Items Saved.";
                    }
                    SSUtils.EndExcelOperation(app, msg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnSaveSelected_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "Program");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = SSUtils.GetListOfProjects(app);
                    bool multiple = JiraShared.ExecuteSaveSelectedCellsToJira(jira, app, listofProjects);
                    string msg = string.Empty;
                    if (multiple == true)
                    {
                        msg = "Selected Items Saved.";
                    }
                    SSUtils.EndExcelOperation(app, msg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // STANDARDIZE TABLES
        private void btnStandardizeTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                TableStandardization.Execute(app, TableStandardization.StandardizationType.Thorough);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnResetView_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                TableStandardization.Execute(app, TableStandardization.StandardizationType.Light);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnToggleProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableStandardization.ExecuteToggleProperties(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // DOT-ONLY
        private void btnMailMerge_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "DOT");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    MailMerge.ExecuteMailMerge_DOT(jira, app);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateRoadMap_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                //SSUtils.BeginExcelOperation(app);
                RoadMap.ExecuteUpdateRoadMap(app);
                //SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnEmailStatus_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Email.ExecuteEmailStatus(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // DOT VIEWS
        private void btnViewReleaseNotes_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewReleaseNotes_DOT(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewReleasePlan_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewReleasePlan_DOT(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewRequirementsStatus_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewRequirementsStatus_DOT(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewBlockedIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewBlockedIssues_DOT(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewEpicsEstimateActual_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewEpicsEstimateActual(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewRequirementsErrors_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewRequirementsErrors(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // OTHERS
        private void btnUpdateChecklist_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                List<string> listofProjects = new List<string>();
                listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                //ImportFromJira.ExecuteUpdateChecklist(app, listofProjects);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void bntUpdateProjects_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = GetJira(app, "Program");
                if (jira != null)
                {
                    SSUtils.BeginExcelOperation(app);
                    JiraProjects.ExecuteUpdateTable(jira, app);
                    SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
