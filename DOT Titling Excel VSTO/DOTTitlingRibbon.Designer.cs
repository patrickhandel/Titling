using System;
using Microsoft.Office.Tools.Ribbon;

namespace DOT_Titling_Excel_VSTO
{
    partial class DOTTitlingRibbon : RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DOTTitlingRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DOTTitlingRibbon));
            
            //TABS
            this.tabHome = this.Factory.CreateRibbonTab();
            this.tabDOT = this.Factory.CreateRibbonTab();
            this.tabPM = this.Factory.CreateRibbonTab();

            //GROUPS
            this.grpDOT = this.Factory.CreateRibbonGroup();
            this.grpPM = this.Factory.CreateRibbonGroup();

            //BUTTONS
            this.btnUpdate_DOT = this.Factory.CreateRibbonButton();
            this.btnUpdate_Program = this.Factory.CreateRibbonButton();

            this.btnUpdateSelected_DOT = this.Factory.CreateRibbonButton();
            this.btnUpdateSelected_Program = this.Factory.CreateRibbonButton();

            this.btnAdd_DOT = this.Factory.CreateRibbonButton();
            this.btnAdd_Program = this.Factory.CreateRibbonButton();

            this.btnStandardizeTable_DOT = this.Factory.CreateRibbonButton();
            this.btnStandardizeTable_PM = this.Factory.CreateRibbonButton();

            this.btnSaveSelected_DOT = this.Factory.CreateRibbonButton();
            this.btnUpdateEpics_DOT = this.Factory.CreateRibbonButton();
            this.bntUpdateProjects = this.Factory.CreateRibbonButton();
            this.btnUpdateChecklist = this.Factory.CreateRibbonButton();
            this.btnUpdateRoadMap_DOT = this.Factory.CreateRibbonButton();
            this.btnMailMerge_DOT = this.Factory.CreateRibbonButton();
            this.btnUpdateTicketDeveloper_DOT = this.Factory.CreateRibbonButton();

            this.btnEmailStatus_DOT = this.Factory.CreateRibbonButton();
            this.btnResetView = this.Factory.CreateRibbonButton();

            this.Views_DOT = this.Factory.CreateRibbonGallery();
            this.btnViewReleasePlan_DOT = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsErrors_DOT = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsStatus_DOT = this.Factory.CreateRibbonButton();
            this.btnViewBlockedTickets_DOT = this.Factory.CreateRibbonButton();
            this.btnToggleProperties = this.Factory.CreateRibbonButton();

            // SUSPEND LAYOUTS
            this.tabHome.SuspendLayout();
            this.grpDOT.SuspendLayout();
            this.tabDOT.SuspendLayout();
            this.tabPM.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHome
            // 
            this.tabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabHome.ControlId.OfficeId = "TabHome";
            this.tabHome.Label = "TabHome";
            this.tabHome.Name = "tabHome";
            // 
            // grpDOT
            // 
            this.grpDOT.Items.Add(this.btnUpdate_DOT);
            this.grpDOT.Items.Add(this.btnUpdateSelected_DOT);
            this.grpDOT.Items.Add(this.btnAdd_DOT);
            this.grpDOT.Items.Add(this.btnSaveSelected_DOT);
            this.grpDOT.Items.Add(this.btnUpdateEpics_DOT);
            this.grpDOT.Items.Add(this.bntUpdateProjects);
            this.grpDOT.Items.Add(this.btnUpdateChecklist);
            this.grpDOT.Items.Add(this.btnUpdateRoadMap_DOT);
            this.grpDOT.Items.Add(this.btnMailMerge_DOT);
            this.grpDOT.Items.Add(this.btnUpdateTicketDeveloper_DOT);
            this.grpDOT.Items.Add(this.btnStandardizeTable_DOT);
            this.grpDOT.Items.Add(this.btnEmailStatus_DOT);
            this.grpDOT.Items.Add(this.btnResetView);
            this.grpDOT.Items.Add(this.Views_DOT);
            this.grpDOT.Label = "DOT Titling";
            this.grpDOT.Name = "grpDOT";
            // 
            // grpPM
            // 
            this.grpPM.Name = "grpPM";
            this.grpPM.Items.Add(this.btnUpdate_Program);
            this.grpPM.Items.Add(this.btnUpdateSelected_Program);
            this.grpPM.Items.Add(this.btnAdd_Program);
            this.grpPM.Items.Add(this.btnStandardizeTable_PM);
            // 
            // tabDOT
            // 
            this.tabDOT.Groups.Add(this.grpDOT);
            this.tabDOT.Label = "DOT";
            this.tabDOT.Name = "tabDOT";
            // 
            // tabPM
            // 
            this.tabPM.Groups.Add(this.grpPM);
            this.tabPM.Label = "WIN PM";
            this.tabPM.Name = "tabPM";
            // 
            // btnUpdate_DOT
            // 
            this.btnUpdate_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Update.Image")));
            this.btnUpdate_DOT.Label = "Update All Tickets";
            this.btnUpdate_DOT.Name = "btnUpdate_DOT";
            this.btnUpdate_DOT.ShowImage = true;
            this.btnUpdate_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_DOT_Click);
            // 
            // btnUpdate_Program
            // 
            this.btnUpdate_Program.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate_Program.Image = ((System.Drawing.Image)(resources.GetObject("Update.Image")));
            this.btnUpdate_Program.Label = "Update All Tickets";
            this.btnUpdate_Program.Name = "btnUpdate_Program";
            this.btnUpdate_Program.ShowImage = true;
            this.btnUpdate_Program.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Program_Click);
            // 
            // btnUpdateSelected_DOT
            // 
            this.btnUpdateSelected_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateSelected_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Update.Image")));
            this.btnUpdateSelected_DOT.Label = "Update Selected Tickets";
            this.btnUpdateSelected_DOT.Name = "btnUpdateSelected_DOT";
            this.btnUpdateSelected_DOT.ShowImage = true;
            this.btnUpdateSelected_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateSelected_DOT_Click);
            // 
            // btnUpdateSelected_Program
            // 
            this.btnUpdateSelected_Program.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateSelected_Program.Image = ((System.Drawing.Image)(resources.GetObject("Update.Image")));
            this.btnUpdateSelected_Program.Label = "Update Selected Tickets";
            this.btnUpdateSelected_Program.Name = "btnUpdateSelected_Program";
            this.btnUpdateSelected_Program.ShowImage = true;
            this.btnUpdateSelected_Program.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateSelected_DOT_Click);
            // 
            // btnAdd_DOT
            // 
            this.btnAdd_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAdd_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Add.Image")));
            this.btnAdd_DOT.Label = "Add New Tickets";
            this.btnAdd_DOT.Name = "btnAdd_DOT";
            this.btnAdd_DOT.ShowImage = true;
            this.btnAdd_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdd_DOT_Click);
            // 
            // btnAdd_Program
            // 
            this.btnAdd_Program.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAdd_Program.Image = ((System.Drawing.Image)(resources.GetObject("Add.Image")));
            this.btnAdd_Program.Label = "Add New Tickets";
            this.btnAdd_Program.Name = "btnAdd_Program";
            this.btnAdd_Program.ShowImage = true;
            this.btnAdd_Program.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdd_Progam_Click);
            // 
            // btnSaveSelected_DOT
            // 
            this.btnSaveSelected_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveSelected_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Save.Image")));
            this.btnSaveSelected_DOT.Label = "Save Selected";
            this.btnSaveSelected_DOT.Name = "btnSaveSelected_DOT";
            this.btnSaveSelected_DOT.ShowImage = true;
            this.btnSaveSelected_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSelected_DOT_Click);
            // 
            // btnUpdateEpics_DOT
            // 
            this.btnUpdateEpics_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateEpics_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Epics.Image")));
            this.btnUpdateEpics_DOT.Label = "Update Epics";
            this.btnUpdateEpics_DOT.Name = "btnUpdateEpics_DOT";
            this.btnUpdateEpics_DOT.ShowImage = true;
            this.btnUpdateEpics_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateEpics_DOT_Click);
            // 
            // bntUpdateProjects
            // 
            this.bntUpdateProjects.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.bntUpdateProjects.Image = ((System.Drawing.Image)(resources.GetObject("Projects.Image")));
            this.bntUpdateProjects.Label = "Update Projects";
            this.bntUpdateProjects.Name = "bntUpdateProjects";
            this.bntUpdateProjects.ShowImage = true;
            this.bntUpdateProjects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bntUpdateProjects_Click);
            // 
            // btnUpdateChecklist
            // 
            this.btnUpdateChecklist.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateChecklist.Image = ((System.Drawing.Image)(resources.GetObject("Checklist.Image")));
            this.btnUpdateChecklist.Label = "Update Checklist";
            this.btnUpdateChecklist.Name = "btnUpdateChecklist";
            this.btnUpdateChecklist.ShowImage = true;
            this.btnUpdateChecklist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateChecklist_Click);
            // 
            // btnUpdateRoadMap_DOT
            // 
            this.btnUpdateRoadMap_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateRoadMap_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Roadmap.Image")));
            this.btnUpdateRoadMap_DOT.Label = "Update Roadmap";
            this.btnUpdateRoadMap_DOT.Name = "btnUpdateRoadMap_DOT";
            this.btnUpdateRoadMap_DOT.ShowImage = true;
            this.btnUpdateRoadMap_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRoadMap_DOT_Click);
            // 
            // btnMailMerge_DOT
            // 
            this.btnMailMerge_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMailMerge_DOT.Image = ((System.Drawing.Image)(resources.GetObject("MailMerge.Image")));
            this.btnMailMerge_DOT.Label = "Mail Merge";
            this.btnMailMerge_DOT.Name = "btnMailMerge_DOT";
            this.btnMailMerge_DOT.ShowImage = true;
            this.btnMailMerge_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMailMerge_DOT_Click);
            // 
            // btnUpdateTicketDeveloper_DOT
            // 
            this.btnUpdateTicketDeveloper_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateTicketDeveloper_DOT.Image = ((System.Drawing.Image)(resources.GetObject("History.Image")));
            this.btnUpdateTicketDeveloper_DOT.Label = "Get History";
            this.btnUpdateTicketDeveloper_DOT.Name = "btnUpdateDeveloperFromHistory";
            this.btnUpdateTicketDeveloper_DOT.ShowImage = true;
            this.btnUpdateTicketDeveloper_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateDeveloper_DOT_Click);
            // 
            // btnStandardizeTable_DOT
            // 
            this.btnStandardizeTable_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStandardizeTable_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Standardize.Image")));
            this.btnStandardizeTable_DOT.Label = "Standardize Table";
            this.btnStandardizeTable_DOT.Name = "btnStandardizeTable_DOT";
            this.btnStandardizeTable_DOT.ShowImage = true;
            this.btnStandardizeTable_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStandardizeTable_Click);
            // 
            // btnStandardizeTable_PM
            // 
            this.btnStandardizeTable_PM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStandardizeTable_PM.Image = ((System.Drawing.Image)(resources.GetObject("Standardize.Image")));
            this.btnStandardizeTable_PM.Label = "Standardize Table";
            this.btnStandardizeTable_PM.Name = "btnStandardizeTable_PM";
            this.btnStandardizeTable_PM.ShowImage = true;
            this.btnStandardizeTable_PM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStandardizeTable_Click);
            // 
            // btnEmailStatus_DOT
            // 
            this.btnEmailStatus_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEmailStatus_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Email.Image")));
            this.btnEmailStatus_DOT.Label = "Email Status";
            this.btnEmailStatus_DOT.Name = "btnEmailStatus_DOT";
            this.btnEmailStatus_DOT.ShowImage = true;
            this.btnEmailStatus_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmailStatus_DOT_Click);
            // 
            // btnResetView
            // 
            this.btnResetView.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnResetView.Image = ((System.Drawing.Image)(resources.GetObject("Reset.Image")));
            this.btnResetView.Label = "Reset View";
            this.btnResetView.Name = "btnResetView";
            this.btnResetView.ShowImage = true;
            this.btnResetView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetView_Click);
            // 
            // Views_DOT
            // 
            this.Views_DOT.Buttons.Add(this.btnViewReleasePlan_DOT);
            this.Views_DOT.Buttons.Add(this.btnViewRequirementsErrors_DOT);
            this.Views_DOT.Buttons.Add(this.btnViewRequirementsStatus_DOT);
            this.Views_DOT.Buttons.Add(this.btnViewBlockedTickets_DOT);
            this.Views_DOT.Buttons.Add(this.btnToggleProperties);
            this.Views_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Views_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.Views_DOT.Label = "Views";
            this.Views_DOT.Name = "Views_DOT";
            this.Views_DOT.ShowImage = true;
            this.Views_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Views_DOT_Click);
            // 
            // btnViewReleasePlan_DOT
            // 
            this.btnViewReleasePlan_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewReleasePlan_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.btnViewReleasePlan_DOT.Description = "Release Schedule";
            this.btnViewReleasePlan_DOT.Label = "Release Schedule";
            this.btnViewReleasePlan_DOT.Name = "btnViewReleasePlan_DOT";
            this.btnViewReleasePlan_DOT.ShowImage = true;
            this.btnViewReleasePlan_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewReleasePlan_DOT_Click);
            // 
            // btnViewRequirementsErrors_DOT
            // 
            this.btnViewRequirementsErrors_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewRequirementsErrors_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.btnViewRequirementsErrors_DOT.Description = "Requirements Errors";
            this.btnViewRequirementsErrors_DOT.Label = "Requirements Errors";
            this.btnViewRequirementsErrors_DOT.Name = "btnViewRequirementsErrors_DOT";
            this.btnViewRequirementsErrors_DOT.ShowImage = true;
            this.btnViewRequirementsErrors_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewRequirementsErrors_DOT_Click);
            // 
            // btnViewRequirementsStatus_DOT
            // 
            this.btnViewRequirementsStatus_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewRequirementsStatus_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.btnViewRequirementsStatus_DOT.Description = "Requirements Status";
            this.btnViewRequirementsStatus_DOT.Label = "Requirements Status";
            this.btnViewRequirementsStatus_DOT.Name = "btnViewRequirementsStatus_DOT";
            this.btnViewRequirementsStatus_DOT.ShowImage = true;
            this.btnViewRequirementsStatus_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewRequirementsStatus_DOT_Click);
            // 
            // btnViewBlockedTickets_DOT
            // 
            this.btnViewBlockedTickets_DOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewBlockedTickets_DOT.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.btnViewBlockedTickets_DOT.Description = "Blocked Tickets";
            this.btnViewBlockedTickets_DOT.Label = "Blocked Tickets";
            this.btnViewBlockedTickets_DOT.Name = "btnViewBlockedTickets_DOT";
            this.btnViewBlockedTickets_DOT.ShowImage = true;
            this.btnViewBlockedTickets_DOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewBlockedTickets_DOT_Click);
            // 
            // btnToggleProperties
            // 
            this.btnToggleProperties.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnToggleProperties.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.btnToggleProperties.Description = "Toggle Properties";
            this.btnToggleProperties.Label = "Toggle Properties";
            this.btnToggleProperties.Name = "btnToggleProperties";
            this.btnToggleProperties.ShowImage = true;
            this.btnToggleProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleProperties_Click);
            // 
            // DOTTitlingRibbon
            // 
            this.Name = "DOTTitlingRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabHome);
            this.Tabs.Add(this.tabDOT);
            this.Tabs.Add(this.tabPM);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DOTTitlingRibbon_Load);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.grpDOT.ResumeLayout(false);
            this.grpDOT.PerformLayout();
            this.tabDOT.ResumeLayout(false);
            this.tabDOT.PerformLayout();
            this.tabPM.ResumeLayout(false);
            this.tabPM.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        internal RibbonTab tabHome;
        internal RibbonTab tabDOT;
        internal RibbonTab tabPM;
        //
        internal RibbonGroup grpDOT;
        internal RibbonGroup grpPM;
        //
        internal RibbonButton btnStandardizeTable_DOT;
        internal RibbonButton btnStandardizeTable_PM;
        internal RibbonButton btnUpdate_DOT;
        internal RibbonButton btnUpdate_Program;
        internal RibbonButton btnUpdateSelected_DOT;
        internal RibbonButton btnUpdateSelected_Program;
        internal RibbonButton btnAdd_DOT;
        internal RibbonButton btnAdd_Program;
        //
        internal RibbonButton btnMailMerge_DOT;
        internal RibbonButton btnResetView;
        internal RibbonButton btnUpdateChecklist;
        internal RibbonButton bntUpdateProjects;
        internal RibbonButton btnSaveSelected_DOT;
        internal RibbonButton btnUpdateRoadMap_DOT;
        internal RibbonButton btnUpdateEpics_DOT;
        internal RibbonButton btnUpdateTicketDeveloper_DOT;
        internal RibbonButton btnEmailStatus_DOT;
        //
        internal RibbonGallery Views_DOT;
            private RibbonButton btnViewReleasePlan_DOT;
            private RibbonButton btnViewRequirementsErrors_DOT;
            private RibbonButton btnViewRequirementsStatus_DOT;
            private RibbonButton btnViewBlockedTickets_DOT;
            private RibbonButton btnToggleProperties;
        //


    }

    partial class ThisRibbonCollection
    {
        internal DOTTitlingRibbon DOTTitlingRibbon
        {
            get { return this.GetRibbon<DOTTitlingRibbon>(); }
        }
    }
}
