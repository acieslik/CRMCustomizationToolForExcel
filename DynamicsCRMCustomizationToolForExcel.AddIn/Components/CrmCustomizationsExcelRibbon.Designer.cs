namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    partial class CrmCustomizationsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CrmCustomizationsRibbon(): base(Globals.Factory.GetRibbonFactory())
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
            this.grpCrmCustomizationsRibbon = this.Factory.CreateRibbonTab();
            this.grp_CrmOperation = this.Factory.CreateRibbonGroup();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.btnPublishAll = this.Factory.CreateRibbonButton();
            this.btnExportChanges = this.Factory.CreateRibbonButton();
            this.btnUpdateSheet = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnGetEntitiesList = this.Factory.CreateRibbonButton();
            this.btnSynchronizeEntity = this.Factory.CreateRibbonButton();
            this.grpCrmCustomizationsRibbon.SuspendLayout();
            this.grp_CrmOperation.SuspendLayout();
            this.group3.SuspendLayout();
            // 
            // grpCrmCustomizationsRibbon
            // 
            this.grpCrmCustomizationsRibbon.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.grpCrmCustomizationsRibbon.Groups.Add(this.grp_CrmOperation);
            this.grpCrmCustomizationsRibbon.Groups.Add(this.group3);
            this.grpCrmCustomizationsRibbon.Label = "CRM Customizations";
            this.grpCrmCustomizationsRibbon.Name = "grpCrmCustomizationsRibbon";
            // 
            // grp_CrmOperation
            // 
            this.grp_CrmOperation.Items.Add(this.btnConnect);
            this.grp_CrmOperation.Items.Add(this.btnExportChanges);
            this.grp_CrmOperation.Items.Add(this.btnUpdateSheet);
            this.grp_CrmOperation.Items.Add(this.btnPublishAll);
            this.grp_CrmOperation.Label = "Crm Operation";
            this.grp_CrmOperation.Name = "grp_CrmOperation";
            // 
            // btnConnect
            // 
            this.btnConnect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConnect.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.Connect;
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.ShowImage = true;
            this.btnConnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConnect_Click);
            // 
            // btnPublishAll
            // 
            this.btnPublishAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPublishAll.Enabled = false;
            this.btnPublishAll.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.PublishAll_32;
            this.btnPublishAll.Label = "Publish All";
            this.btnPublishAll.Name = "btnPublishAll";
            this.btnPublishAll.ShowImage = true;
            this.btnPublishAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPublishAll_Click);
            // 
            // btnExportChanges
            // 
            this.btnExportChanges.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportChanges.Enabled = false;
            this.btnExportChanges.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.ImportData32;
            this.btnExportChanges.Label = "Send Changes  To Crm";
            this.btnExportChanges.Name = "btnExportChanges";
            this.btnExportChanges.ShowImage = true;
            this.btnExportChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportChanges_Click);
            // 
            // btnUpdateSheet
            // 
            this.btnUpdateSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateSheet.Enabled = false;
            this.btnUpdateSheet.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.Refresh_32;
            this.btnUpdateSheet.Label = "Update Sheet";
            this.btnUpdateSheet.Name = "btnUpdateSheet";
            this.btnUpdateSheet.ShowImage = true;
            this.btnUpdateSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateSheet_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnGetEntitiesList);
            this.group3.Items.Add(this.btnSynchronizeEntity);
            this.group3.Label = "Crm Enitities Operations";
            this.group3.Name = "group3";
            this.group3.Visible = false;
            // 
            // btnGetEntitiesList
            // 
            this.btnGetEntitiesList.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetEntitiesList.Enabled = false;
            this.btnGetEntitiesList.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.ExportToExcel32;
            this.btnGetEntitiesList.Label = "Get Entities List";
            this.btnGetEntitiesList.Name = "btnGetEntitiesList";
            this.btnGetEntitiesList.ShowImage = true;
            this.btnGetEntitiesList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetEntitiesList_Click);
            // 
            // btnSynchronizeEntity
            // 
            this.btnSynchronizeEntity.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSynchronizeEntity.Enabled = false;
            this.btnSynchronizeEntity.Image = global::DynamicsCRMCustomizationToolForExcel.AddIn.Properties.Resources.ConvertToKit_32;
            this.btnSynchronizeEntity.Label = "Synchronize Entity";
            this.btnSynchronizeEntity.Name = "btnSynchronizeEntity";
            this.btnSynchronizeEntity.ShowImage = true;
            this.btnSynchronizeEntity.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSynchronizeEntity_Click);
            // 
            // CrmCustomizationsRibbon
            // 
            this.Name = "CrmCustomizationsRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.grpCrmCustomizationsRibbon);
            this.grpCrmCustomizationsRibbon.ResumeLayout(false);
            this.grpCrmCustomizationsRibbon.PerformLayout();
            this.grp_CrmOperation.ResumeLayout(false);
            this.grp_CrmOperation.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab grpCrmCustomizationsRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_CrmOperation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportChanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPublishAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetEntitiesList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSynchronizeEntity;
    }

    partial class ThisRibbonCollection
    {
        internal CrmCustomizationsRibbon Ribbon1
        {
            get { return this.GetRibbon<CrmCustomizationsRibbon>(); }
        }
    }
}
