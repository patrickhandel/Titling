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

        private void btnUpdateRoadMap_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                //SSUtils.BeginExcelOperation(app);
                RoadMap.ExecuteUpdateRoadMap();
                //SSUtils.EndExcelOperation(app, string.Empty);
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
                SSUtils.EndExcelOperation(app, string.Empty);
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
                SSUtils.EndExcelOperation(app, "Selected Items Updated");
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
                SSUtils.EndExcelOperation(app, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            //BgWorker.Form1.EnableVisualStyles();
            //app.SetCompatibleTextRenderingDefault(false);
            //app.Run(new BgWorker.Form1());
            //Form2.ActiveForm;

            //Form2 frm = new Form2();

            //frm.Activate();

            app.Run(new Form2());

        }

        private void btnImportEpics_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ImportFromJira.ExecuteUpdateEpics();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
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



        private void btnDeveloperFromHistory_Click(object sender, RibbonControlEventArgs e)
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

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = app.ActiveSheet;
            SSUtils.GetListOfTables(activeWorksheet);
        }
    }
}
