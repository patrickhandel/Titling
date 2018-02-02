﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

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
        private void btnUpdate_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = false;
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteUpdateTickets(app, forProgram, projectKey);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdate_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = true;
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteUpdateTickets(app, forProgram, projectKey);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAdd_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = false;
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteAdd_DOT(app, forProgram, projectKey);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAdd_Progam_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                //bool forProgram = true;
                //ImportFromJira.ExecuteAdd_Program(app, forProgram);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateSelected_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteUpdateSelected_DOT(app, projectKey);
                SSUtils.EndExcelOperation(app, "Selected Items Updated");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateSelected_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                //ImportFromJira.ExecuteUpdateSelected_Program(app);
                SSUtils.EndExcelOperation(app, "Selected Items Updated");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnSaveSelected_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool multiple = ExportToJira.SaveSelected_DOT(app);
                string msg = string.Empty;
                if (multiple == true)
                {
                    msg = "Selected Items Saved.";
                }
                SSUtils.EndExcelOperation(app, msg);
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
                SSUtils.BeginExcelOperation(app);
                //bool multiple = ExportToJira.SaveSelected_Program(app);
                //string msg = string.Empty;
                //if (multiple == true)
                //{
                //    msg = "Selected Items Saved.";
                //}
                //SSUtils.EndExcelOperation(app, msg);
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
                TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Thorough);
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
                TableStandardization.ExecuteStandardizeTable(app, TableStandardization.StandardizationType.Light);
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
                SSUtils.BeginExcelOperation(app);
                MailMerge.ExecuteMailMerge_DOT(app);
                SSUtils.EndExcelOperation(app, string.Empty);
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
                RoadMap.ExecuteUpdateRoadMap_DOT(app);
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
                Email.ExecuteEmailStatus_DOT(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateEpics_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteUpdateEpics_DOT(app, projectKey);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateDeveloper_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                History.ExecuteUpdateDeveloper_DOT();
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        // DOT VIEWS
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

        private void btnViewBlockedTickets_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewBlockedTickets_DOT(app);
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
                string projectKey = ThisAddIn.ProjectKeyDOT;
                ImportFromJira.ExecuteUpdateChecklist(app, projectKey);
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
                SSUtils.BeginExcelOperation(app);
                //ImportFromJira.ExecuteAddNewProjects(app);
                ImportFromJira.ExecuteUpdateProjects(app);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
