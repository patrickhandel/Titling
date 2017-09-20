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

        private void btnCleanup_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = app.ActiveSheet;
            WorksheetStandardization.ExecuteCleanup(activeWorksheet);
        }

        private void btnAddNewTickets_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Not Implemented");
            ImportFromJira.ExecuteAddNewTickets();
        }

        private void btnImportSelected_Click(object sender, RibbonControlEventArgs e)
        {
            ImportFromJira.ExecuteImportSelectedJiraTickets();
        }

        private void btnImportAll_Click(object sender, RibbonControlEventArgs e)
        {
            //Import.ExecuteImportAllJiraTickets();
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            ExportToJira.ExecuteSaveSummary();
        }
    }
}
