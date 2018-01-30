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
            this.tabHome = this.Factory.CreateRibbonTab();
            this.grpJira = this.Factory.CreateRibbonGroup();
            this.btnImportAll = this.Factory.CreateRibbonButton();
            this.btnImportSelected = this.Factory.CreateRibbonButton();
            this.btnAddNewTickets = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnImportEpics = this.Factory.CreateRibbonButton();
            this.btnImportProjects = this.Factory.CreateRibbonButton();
            this.btnImportChecklist = this.Factory.CreateRibbonButton();
            this.btnUpdateRoadMap = this.Factory.CreateRibbonButton();
            this.btnMailMerge = this.Factory.CreateRibbonButton();
            this.btnDeveloperFromHistory = this.Factory.CreateRibbonButton();
            this.btnCleanup = this.Factory.CreateRibbonButton();
            this.btnEmail = this.Factory.CreateRibbonButton();
            this.btnResetView = this.Factory.CreateRibbonButton();
            this.Views = this.Factory.CreateRibbonGallery();
            this.btnViewReleasePlan = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsErrors = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsStatus = this.Factory.CreateRibbonButton();
            this.btnViewBlockedTickets = this.Factory.CreateRibbonButton();
            this.btnShowHidePropertiesRow = this.Factory.CreateRibbonButton();
            this.tabJira = this.Factory.CreateRibbonTab();
            this.tabHome.SuspendLayout();
            this.grpJira.SuspendLayout();
            this.tabJira.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHome
            // 
            this.tabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabHome.ControlId.OfficeId = "TabHome";
            this.tabHome.Label = "TabHome";
            this.tabHome.Name = "tabHome";
            // 
            // grpJira
            // 
            this.grpJira.Items.Add(this.btnImportAll);
            this.grpJira.Items.Add(this.btnImportSelected);
            this.grpJira.Items.Add(this.btnAddNewTickets);
            this.grpJira.Items.Add(this.btnUpdate);
            this.grpJira.Items.Add(this.btnImportEpics);
            this.grpJira.Items.Add(this.btnImportProjects);
            this.grpJira.Items.Add(this.btnImportChecklist);
            this.grpJira.Items.Add(this.btnUpdateRoadMap);
            this.grpJira.Items.Add(this.btnMailMerge);
            this.grpJira.Items.Add(this.btnDeveloperFromHistory);
            this.grpJira.Items.Add(this.btnCleanup);
            this.grpJira.Items.Add(this.btnEmail);
            this.grpJira.Items.Add(this.btnResetView);
            this.grpJira.Items.Add(this.Views);
            this.grpJira.Label = "DOT Titling";
            this.grpJira.Name = "grpJira";
            // 
            // btnImportAll
            // 
            this.btnImportAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportAll.Image = ((System.Drawing.Image)(resources.GetObject("btnImportAll.Image")));
            this.btnImportAll.Label = "Update All Tickets";
            this.btnImportAll.Name = "btnImportAll";
            this.btnImportAll.ShowImage = true;
            this.btnImportAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportAllTickets_Click);
            // 
            // btnImportSelected
            // 
            this.btnImportSelected.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportSelected.Image = ((System.Drawing.Image)(resources.GetObject("btnImportSelected.Image")));
            this.btnImportSelected.Label = "Update Selected Tickets";
            this.btnImportSelected.Name = "btnImportSelected";
            this.btnImportSelected.ShowImage = true;
            this.btnImportSelected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportSelectedTickets_Click);
            // 
            // btnAddNewTickets
            // 
            this.btnAddNewTickets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddNewTickets.Image = ((System.Drawing.Image)(resources.GetObject("btnAddNewTickets.Image")));
            this.btnAddNewTickets.Label = "Add New Tickets";
            this.btnAddNewTickets.Name = "btnAddNewTickets";
            this.btnAddNewTickets.ShowImage = true;
            this.btnAddNewTickets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddNewTickets_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdate.Image")));
            this.btnUpdate.Label = "Save Selected";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // btnImportEpics
            // 
            this.btnImportEpics.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportEpics.Image = ((System.Drawing.Image)(resources.GetObject("btnImportEpics.Image")));
            this.btnImportEpics.Label = "Update Epics";
            this.btnImportEpics.Name = "btnImportEpics";
            this.btnImportEpics.ShowImage = true;
            this.btnImportEpics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportEpics_Click);
            // 
            // btnImportProjects
            // 
            this.btnImportProjects.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportProjects.Image = global::DOT_Titling_Excel_VSTO.Properties.Resources.project;
            this.btnImportProjects.Label = "Update Projects";
            this.btnImportProjects.Name = "btnImportProjects";
            this.btnImportProjects.ShowImage = true;
            this.btnImportProjects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportProjects_Click);
            // 
            // btnImportChecklist
            // 
            this.btnImportChecklist.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportChecklist.Image = global::DOT_Titling_Excel_VSTO.Properties.Resources.To_Do_List_512;
            this.btnImportChecklist.Label = "Update Checklist";
            this.btnImportChecklist.Name = "btnImportChecklist";
            this.btnImportChecklist.ShowImage = true;
            this.btnImportChecklist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportChecklist_Click);
            // 
            // btnUpdateRoadMap
            // 
            this.btnUpdateRoadMap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateRoadMap.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateRoadMap.Image")));
            this.btnUpdateRoadMap.Label = "Update Roadmap";
            this.btnUpdateRoadMap.Name = "btnUpdateRoadMap";
            this.btnUpdateRoadMap.ShowImage = true;
            this.btnUpdateRoadMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRoadMap_Click);
            // 
            // btnMailMerge
            // 
            this.btnMailMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMailMerge.Image = ((System.Drawing.Image)(resources.GetObject("btnMailMerge.Image")));
            this.btnMailMerge.Label = "Mail Merge";
            this.btnMailMerge.Name = "btnMailMerge";
            this.btnMailMerge.ShowImage = true;
            this.btnMailMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMailMerge_Click);
            // 
            // btnDeveloperFromHistory
            // 
            this.btnDeveloperFromHistory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeveloperFromHistory.Image = ((System.Drawing.Image)(resources.GetObject("btnDeveloperFromHistory.Image")));
            this.btnDeveloperFromHistory.Label = "Get History";
            this.btnDeveloperFromHistory.Name = "btnDeveloperFromHistory";
            this.btnDeveloperFromHistory.ShowImage = true;
            this.btnDeveloperFromHistory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeveloperFromHistory_Click);
            // 
            // btnCleanup
            // 
            this.btnCleanup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCleanup.Image = ((System.Drawing.Image)(resources.GetObject("btnCleanup.Image")));
            this.btnCleanup.Label = "Cleanup Worksheet";
            this.btnCleanup.Name = "btnCleanup";
            this.btnCleanup.ShowImage = true;
            this.btnCleanup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanupTable_Click);
            // 
            // btnEmail
            // 
            this.btnEmail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEmail.Image = global::DOT_Titling_Excel_VSTO.Properties.Resources.email_2_icon;
            this.btnEmail.Label = "Email Status";
            this.btnEmail.Name = "btnEmail";
            this.btnEmail.ShowImage = true;
            this.btnEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmail_Click);
            // 
            // btnResetView
            // 
            this.btnResetView.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnResetView.Image = ((System.Drawing.Image)(resources.GetObject("btnResetView.Image")));
            this.btnResetView.Label = "Reset View";
            this.btnResetView.Name = "btnResetView";
            this.btnResetView.ShowImage = true;
            this.btnResetView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetView_Click);
            // 
            // Views
            // 
            this.Views.Buttons.Add(this.btnViewReleasePlan);
            this.Views.Buttons.Add(this.btnViewRequirementsErrors);
            this.Views.Buttons.Add(this.btnViewRequirementsStatus);
            this.Views.Buttons.Add(this.btnViewBlockedTickets);
            this.Views.Buttons.Add(this.btnShowHidePropertiesRow);
            this.Views.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Views.Image = ((System.Drawing.Image)(resources.GetObject("Views.Image")));
            this.Views.Label = "Views";
            this.Views.Name = "Views";
            this.Views.ShowImage = true;
            this.Views.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Views_Click);
            // 
            // btnViewReleasePlan
            // 
            this.btnViewReleasePlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewReleasePlan.Description = "Release Schedule";
            this.btnViewReleasePlan.Label = "Release Schedule";
            this.btnViewReleasePlan.Name = "btnViewReleasePlan";
            this.btnViewReleasePlan.ShowImage = true;
            this.btnViewReleasePlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewReleasePlan_Click);
            // 
            // btnViewRequirementsErrors
            // 
            this.btnViewRequirementsErrors.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewRequirementsErrors.Description = "Requirements Errors";
            this.btnViewRequirementsErrors.Label = "Requirements Errors";
            this.btnViewRequirementsErrors.Name = "btnViewRequirementsErrors";
            this.btnViewRequirementsErrors.ShowImage = true;
            this.btnViewRequirementsErrors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewRequirementsErrors_Click);
            // 
            // btnViewRequirementsStatus
            // 
            this.btnViewRequirementsStatus.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewRequirementsStatus.Description = "Requirements Status";
            this.btnViewRequirementsStatus.Label = "Requirements Status";
            this.btnViewRequirementsStatus.Name = "btnViewRequirementsStatus";
            this.btnViewRequirementsStatus.ShowImage = true;
            this.btnViewRequirementsStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewRequirementsStatus_Click);
            // 
            // btnViewBlockedTickets
            // 
            this.btnViewBlockedTickets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewBlockedTickets.Description = "Blocked Tickets";
            this.btnViewBlockedTickets.Label = "Blocked Tickets";
            this.btnViewBlockedTickets.Name = "btnViewBlockedTickets";
            this.btnViewBlockedTickets.ShowImage = true;
            this.btnViewBlockedTickets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewBlockedTickets_Click);
            // 
            // btnShowHidePropertiesRow
            // 
            this.btnShowHidePropertiesRow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnShowHidePropertiesRow.Description = "Toggle Properties Row";
            this.btnShowHidePropertiesRow.Label = "Toggle Properties Row";
            this.btnShowHidePropertiesRow.Name = "btnShowHidePropertiesRow";
            this.btnShowHidePropertiesRow.ShowImage = true;
            this.btnShowHidePropertiesRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowHidePropertiesRow_Click);
            // 
            // tabJira
            // 
            this.tabJira.Groups.Add(this.grpJira);
            this.tabJira.Label = "Jira";
            this.tabJira.Name = "tabJira";
            // 
            // DOTTitlingRibbon
            // 
            this.Name = "DOTTitlingRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabHome);
            this.Tabs.Add(this.tabJira);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DOTTitlingRibbon_Load);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.grpJira.ResumeLayout(false);
            this.grpJira.PerformLayout();
            this.tabJira.ResumeLayout(false);
            this.tabJira.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal RibbonTab tabHome;
        internal RibbonGroup grpJira;
        internal RibbonButton btnMailMerge;
        internal RibbonButton btnCleanup;
        internal RibbonButton btnAddNewTickets;
        internal RibbonButton btnImportAll;
        internal RibbonButton btnImportSelected;
        internal RibbonButton btnResetView;
        internal RibbonButton btnImportChecklist;
        internal RibbonButton btnImportProjects;
        internal RibbonButton btnUpdate;
        internal RibbonButton btnUpdateRoadMap;
        internal RibbonButton btnImportEpics;
        internal RibbonButton btnDeveloperFromHistory;
        internal RibbonGallery Views;
        private RibbonButton btnViewReleasePlan;
        private RibbonButton btnViewRequirementsErrors;
        private RibbonButton btnViewRequirementsStatus;
        private RibbonButton btnViewBlockedTickets;
        private RibbonButton btnShowHidePropertiesRow;
        internal RibbonButton btnEmail;
        internal RibbonTab tabJira;
    }

    partial class ThisRibbonCollection
    {
        internal DOTTitlingRibbon DOTTitlingRibbon
        {
            get { return this.GetRibbon<DOTTitlingRibbon>(); }
        }
    }
}
