using Microsoft.Office.Tools.Ribbon;

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
            WorksheetStandardization.ExecuteCleanup();
        }

        private void btnAddNewStories_Click(object sender, RibbonControlEventArgs e)
        {
            Maintenance.AddNewStories();
        }
    }
}
