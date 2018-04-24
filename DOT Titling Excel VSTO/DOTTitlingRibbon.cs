using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
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

        // UPDATE, ADD, SAVE
        private async void btnUpdateIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    success = await JiraShared.ExecuteUpdateTable(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnUpdateIssues_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = await SSUtils.GetListOfProjects(app);
                    success = await JiraShared.ExecuteUpdateTable(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnAddIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    success = await JiraShared.ExecuteAddNewRowsToTable(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnAddIssues_Progam_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = await SSUtils.GetListOfProjects(app);
                    success = await JiraShared.ExecuteAddNewRowsToTable(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnUpdateSelectedIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    success = await JiraShared.ExecuteUpdateSelectedRows(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, "Selected Items Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnUpdateSelectedIssues_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = await SSUtils.GetListOfProjects(app);
                    success = await JiraShared.ExecuteUpdateSelectedRows(jira, app, listofProjects);
                    success = await SSUtils.EndExcelOperation(app, "Selected Items Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnSaveSelectedIssues_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = new List<string>();
                    listofProjects.Add(ThisAddIn.ProjectKeyDOT);
                    bool multiple = await JiraShared.ExecuteSaveSelectedCellsToJira(jira, app, listofProjects);
                    string msg = string.Empty;
                    if (multiple == true)
                    {
                        msg = "Selected Items Saved.";
                    }
                    success = await SSUtils.EndExcelOperation(app, msg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private async void btnSaveSelected_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    List<string> listofProjects = await SSUtils.GetListOfProjects(app);
                    bool multiple = await JiraShared.ExecuteSaveSelectedCellsToJira(jira, app, listofProjects);
                    string msg = string.Empty;
                    if (multiple == true)
                    {
                        msg = "Selected Items Saved.";
                    }
                    bool sucess = await SSUtils.EndExcelOperation(app, msg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // STANDARDIZE TABLES
        private async void btnStandardizeTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                bool success;
                success = await SSUtils.BeginExcelOperation(app);
                success = await TableStandardization.Execute(app, TableStandardization.StandardizationType.Thorough);
                success = await SSUtils.EndExcelOperation(app, string.Empty);
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
        private async void btnMailMerge_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    success = await MailMerge.ExecuteMailMerge_DOT(jira, app);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
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

        private async void bntUpdateProjects_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Jira.Jira jira = await JiraShared.GetJira(app);
                if (jira != null)
                {
                    bool success;
                    success = await SSUtils.BeginExcelOperation(app);
                    success = await ImportData.ExecuteUpdateTable(jira, app);
                    success = await SSUtils.EndExcelOperation(app, string.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnImportData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                    bool success;
                    //success = await SSUtils.BeginExcelOperation(app);
                    success = Metrics.Import(app);
                    //success = await SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }

        }
    }
}
