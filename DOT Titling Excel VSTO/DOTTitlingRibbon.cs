using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    public partial class DOTTitlingRibbon
    {
        private void btnUpdate_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = false;
                ImportFromJira.ExecuteImportTickets(app, forProgram);
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
                ImportFromJira.ExecuteImportTickets(app, forProgram);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAddTickets_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = false;
                ImportFromJira.ExecuteAddTickets(app, forProgram);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAddTickets_Progam_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                bool forProgram = true;
                ImportFromJira.ExecuteAddTickets(app, forProgram);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAdd_Program_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                ImportFromJira.ExecuteAddTickets(app, false);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }



        private void DOTTitlingRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void Views_DOT_Click(object sender, RibbonControlEventArgs e)
        {

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

        private void btnToggleProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableStandardization.ExecuteTogglePropertiesRow(app);
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
                ImportFromJira.ExecuteUpdateSelectedTickets(app);
                SSUtils.EndExcelOperation(app, "Selected Items Updated");
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
                ImportFromJira.ExecuteUpdateEpics(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdateChecklist_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                ImportFromJira.ExecuteUpdateChecklistTasks(app);
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

        private void btnMailMerge_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                MailMerge.ExecuteMailMerge(app);
                SSUtils.EndExcelOperation(app, string.Empty);
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
                bool multiple = ExportToJira.ExecuteSaveSelectedTicketValues(app);
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

        private void btnUpdateTicketDeveloper_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                History.ExecuteGetDeveloperFromHistory();
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnViewReleasePlan_DOT_DOT(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                TableViews.ExecuteViewReleasePlan(app);
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
                TableViews.ExecuteViewRequirementsStatus(app);
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
                TableViews.ExecuteViewBlockedTickets(app);
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

        private void btnEmailStatus_DOT_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Email.SendStatus(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }


    }
}
