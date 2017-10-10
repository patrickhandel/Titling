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
            this.btnCleanup = this.Factory.CreateRibbonButton();
            this.btnMailMerge = this.Factory.CreateRibbonButton();
            this.btnUpdateRoadMap = this.Factory.CreateRibbonButton();
            this.btnImportEpics = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnCleanup);
            this.group1.Items.Add(this.btnMailMerge);
            this.group1.Items.Add(this.btnUpdateRoadMap);
            this.group1.Items.Add(this.btnImportEpics);
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
            // btnCleanup
            // 
            this.btnCleanup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCleanup.Image = ((System.Drawing.Image)(resources.GetObject("btnCleanup.Image")));
            this.btnCleanup.Label = "Cleanup Worksheet";
            this.btnCleanup.Name = "btnCleanup";
            this.btnCleanup.ShowImage = true;
            this.btnCleanup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanupWorksheet_Click);
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
            // btnUpdateRoadMap
            // 
            this.btnUpdateRoadMap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateRoadMap.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateRoadMap.Image")));
            this.btnUpdateRoadMap.Label = "Update Roadmap";
            this.btnUpdateRoadMap.Name = "btnUpdateRoadMap";
            this.btnUpdateRoadMap.ShowImage = true;
            this.btnUpdateRoadMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateRoadMap_Click);
            // 
            // btnImportEpics
            // 
            this.btnImportEpics.Label = "Update Epics";
            this.btnImportEpics.Name = "btnImportEpics";
            this.btnImportEpics.ShowImage = true;
            this.btnImportEpics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportEpics_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateRoadMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportEpics;
    }

    partial class ThisRibbonCollection
    {
        internal DOTTitlingRibbon DOTTitlingRibbon
        {
            get { return this.GetRibbon<DOTTitlingRibbon>(); }
        }
    }
}
