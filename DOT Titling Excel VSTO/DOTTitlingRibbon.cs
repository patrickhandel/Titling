using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    public partial class DOTTitlingRibbon
    {
        private void DOTTitlingRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnMailMerge_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
            SSUtils.BeginExcelOperation(app);
            MailMerge.ExecuteMailMerge();
            SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnCleanupWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet activeWorksheet = app.ActiveSheet;
                SSUtils.BeginExcelOperation(app);
                WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnAddNewTickets_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                ImportFromJira.ExecuteAddNewTickets();
                SSUtils.EndExcelOperation(app, "Ticket Addition");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnImportSelectedTickets_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                ImportFromJira.ExecuteUpdateSelectedTickets();
                SSUtils.EndExcelOperation(app, "Selected Update");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnImportAllTickets_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                ImportFromJira.ExecuteUpdateAllTickets();
                SSUtils.EndExcelOperation(app, "Complete Update");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                SSUtils.BeginExcelOperation(app);
                ExportToJira.ExecuteSaveTicket();
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
