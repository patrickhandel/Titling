using System;
using Microsoft.Office.Tools.Ribbon;

namespace DOT_Titling_Excel_VSTO
{
    partial class DOTTitlingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
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
            this.btnResetView = this.Factory.CreateRibbonButton();
            this.Views = this.Factory.CreateRibbonGallery();
            this.btnViewReleasePlan = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsErrors = this.Factory.CreateRibbonButton();
            this.btnViewRequirementsStatus = this.Factory.CreateRibbonButton();
            this.btnViewBlockedTickets = this.Factory.CreateRibbonButton();
            this.btnShowHidePropertiesRow = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnImportAll);
            this.group1.Items.Add(this.btnImportSelected);
            this.group1.Items.Add(this.btnAddNewTickets);
            this.group1.Items.Add(this.btnUpdate);
            this.group1.Items.Add(this.btnImportEpics);
            this.group1.Items.Add(this.btnImportProjects);
            this.group1.Items.Add(this.btnImportChecklist);
            this.group1.Items.Add(this.btnUpdateRoadMap);
            this.group1.Items.Add(this.btnMailMerge);
            this.group1.Items.Add(this.btnDeveloperFromHistory);
            this.group1.Items.Add(this.btnCleanup);
            this.group1.Items.Add(this.btnResetView);
            this.group1.Items.Add(this.Views);
            this.group1.Label = "DOT Titling";
            this.group1.Name = "group1";
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
            // DOTTitlingRibbon
            // 
            this.Name = "DOTTitlingRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DOTTitlingRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMailMerge;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNewTickets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetView;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportChecklist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportProjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateRoadMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportEpics;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperFromHistory;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery Views;
        //
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnViewReleasePlan;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnViewRequirementsErrors;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnViewRequirementsStatus;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnViewBlockedTickets;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnShowHidePropertiesRow;

    }

    partial class ThisRibbonCollection
    {
        internal DOTTitlingRibbon DOTTitlingRibbon
        {
            get { return this.GetRibbon<DOTTitlingRibbon>(); }
        }
    }
}
