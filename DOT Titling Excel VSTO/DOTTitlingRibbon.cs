using Microsoft.Office.Tools.Ribbon;
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
            MailMerge.ExecuteMailMerge();
        }

        private void btnCleanupWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = app.ActiveSheet;
            WorksheetStandardization.ExecuteCleanupWorksheet(activeWorksheet);
        }

        private void btnAddNewTickets_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFromJira.ExecuteAddNewTickets();
        }

        private void btnImportSelectedTickets_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFromJira.ExecuteUpdateSelectedTickets();
        }

        private void btnImportAllTickets_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFromJira.ExecuteUpdateAllTickets();
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            ExportToJira.ExecuteSaveSummary();
        }
    }
}
