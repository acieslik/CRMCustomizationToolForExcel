namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    partial class MainPanelEnityList
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.components = new System.ComponentModel.Container();
            this.cmbLanguage = new System.Windows.Forms.ComboBox();
            this.cmbSolution = new System.Windows.Forms.ComboBox();
            this.lblSolution = new System.Windows.Forms.Label();
            this.btnOpenAll = new System.Windows.Forms.Button();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.pnlEntities = new System.Windows.Forms.Panel();
            this.trvCrmComponents = new System.Windows.Forms.TreeView();
            this.chbSheetsProtection = new System.Windows.Forms.CheckBox();
            this.btnAddView = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmAddView = new System.Windows.Forms.ToolStripMenuItem();
            this.ctmShowView = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.showViewAnfFetchXml = new System.Windows.Forms.ToolStripMenuItem();
            this.pnlEntities.SuspendLayout();
            this.btnAddView.SuspendLayout();
            this.ctmShowView.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbLanguage
            // 
            this.cmbLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbLanguage.FormattingEnabled = true;
            this.cmbLanguage.Location = new System.Drawing.Point(22, 60);
            this.cmbLanguage.Name = "cmbLanguage";
            this.cmbLanguage.Size = new System.Drawing.Size(240, 21);
            this.cmbLanguage.TabIndex = 3;
            this.cmbLanguage.SelectedIndexChanged += new System.EventHandler(this.cmbLanguage_SelectedIndexChanged);
            // 
            // cmbSolution
            // 
            this.cmbSolution.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbSolution.FormattingEnabled = true;
            this.cmbSolution.Location = new System.Drawing.Point(22, 16);
            this.cmbSolution.Name = "cmbSolution";
            this.cmbSolution.Size = new System.Drawing.Size(240, 21);
            this.cmbSolution.TabIndex = 4;
            this.cmbSolution.SelectedIndexChanged += new System.EventHandler(this.cmbSolution_SelectedIndexChanged);
            // 
            // lblSolution
            // 
            this.lblSolution.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSolution.AutoSize = true;
            this.lblSolution.Location = new System.Drawing.Point(19, 0);
            this.lblSolution.Name = "lblSolution";
            this.lblSolution.Size = new System.Drawing.Size(45, 13);
            this.lblSolution.TabIndex = 5;
            this.lblSolution.Text = "Solution";
            // 
            // btnOpenAll
            // 
            this.btnOpenAll.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenAll.Location = new System.Drawing.Point(22, 365);
            this.btnOpenAll.Name = "btnOpenAll";
            this.btnOpenAll.Size = new System.Drawing.Size(240, 23);
            this.btnOpenAll.TabIndex = 6;
            this.btnOpenAll.Text = "Open All";
            this.btnOpenAll.UseVisualStyleBackColor = true;
            this.btnOpenAll.Click += new System.EventHandler(this.btnOpenAll_Click);
            // 
            // lblLanguage
            // 
            this.lblLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(19, 40);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(55, 13);
            this.lblLanguage.TabIndex = 7;
            this.lblLanguage.Text = "Langauge";
            // 
            // pnlEntities
            // 
            this.pnlEntities.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlEntities.Controls.Add(this.trvCrmComponents);
            this.pnlEntities.Controls.Add(this.chbSheetsProtection);
            this.pnlEntities.Controls.Add(this.cmbSolution);
            this.pnlEntities.Controls.Add(this.btnOpenAll);
            this.pnlEntities.Controls.Add(this.lblLanguage);
            this.pnlEntities.Controls.Add(this.cmbLanguage);
            this.pnlEntities.Controls.Add(this.lblSolution);
            this.pnlEntities.Location = new System.Drawing.Point(3, 3);
            this.pnlEntities.Name = "pnlEntities";
            this.pnlEntities.Size = new System.Drawing.Size(282, 401);
            this.pnlEntities.TabIndex = 8;
            this.pnlEntities.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlEntities_Paint);
            // 
            // trvCrmComponents
            // 
            this.trvCrmComponents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.trvCrmComponents.Location = new System.Drawing.Point(22, 111);
            this.trvCrmComponents.Name = "trvCrmComponents";
            this.trvCrmComponents.Size = new System.Drawing.Size(240, 248);
            this.trvCrmComponents.TabIndex = 9;
            this.trvCrmComponents.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvCrmComponents_AfterSelect);
            // 
            // chbSheetsProtection
            // 
            this.chbSheetsProtection.AutoSize = true;
            this.chbSheetsProtection.Location = new System.Drawing.Point(22, 87);
            this.chbSheetsProtection.Name = "chbSheetsProtection";
            this.chbSheetsProtection.Size = new System.Drawing.Size(146, 17);
            this.chbSheetsProtection.TabIndex = 8;
            this.chbSheetsProtection.Text = "Enable Sheets Protection";
            this.chbSheetsProtection.UseVisualStyleBackColor = true;
            this.chbSheetsProtection.CheckedChanged += new System.EventHandler(this.chbSheetsProtection_CheckedChanged);
            // 
            // btnAddView
            // 
            this.btnAddView.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmAddView});
            this.btnAddView.Name = "ctmView";
            this.btnAddView.Size = new System.Drawing.Size(152, 26);
            // 
            // tsmAddView
            // 
            this.tsmAddView.Name = "tsmAddView";
            this.tsmAddView.Size = new System.Drawing.Size(151, 22);
            this.tsmAddView.Text = "Add New View";
            this.tsmAddView.Click += new System.EventHandler(this.tsmAddView_Click);
            // 
            // ctmShowView
            // 
            this.ctmShowView.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.showViewAnfFetchXml});
            this.ctmShowView.Name = "ctmShowView";
            this.ctmShowView.Size = new System.Drawing.Size(211, 26);
            // 
            // showViewAnfFetchXml
            // 
            this.showViewAnfFetchXml.Name = "showViewAnfFetchXml";
            this.showViewAnfFetchXml.Size = new System.Drawing.Size(210, 22);
            this.showViewAnfFetchXml.Text = "Show View and Fetch Xml";
            this.showViewAnfFetchXml.Click += new System.EventHandler(this.showViewAnfFetchXml_Click);
            // 
            // MainPanelEnityList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlEntities);
            this.Name = "MainPanelEnityList";
            this.Size = new System.Drawing.Size(288, 404);
            this.pnlEntities.ResumeLayout(false);
            this.pnlEntities.PerformLayout();
            this.btnAddView.ResumeLayout(false);
            this.ctmShowView.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbLanguage;
        private System.Windows.Forms.ComboBox cmbSolution;
        private System.Windows.Forms.Label lblSolution;
        private System.Windows.Forms.Button btnOpenAll;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.Panel pnlEntities;
        private System.Windows.Forms.CheckBox chbSheetsProtection;
        private System.Windows.Forms.TreeView trvCrmComponents;
        private System.Windows.Forms.ContextMenuStrip btnAddView;
        private System.Windows.Forms.ToolStripMenuItem tsmAddView;
        private System.Windows.Forms.ContextMenuStrip ctmShowView;
        private System.Windows.Forms.ToolStripMenuItem showViewAnfFetchXml;
    }

}
