using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using DynamicsCRMCustomizationToolForExcel.Controller;
using Microsoft.Office.Tools;

namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    public partial class DynamicsCRMExcelAddIn
    {
        private CustomTaskPane _MainPanel;
        internal CustomTaskPane MainPanel
        {
            get
            {
                return this._MainPanel;
            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            GlobalOperations.Instance.ExcelOperations = new ExcelDataAccess(Globals.DynamicsCRMExcelAddIn.Application);
            MainPanelEnityList taskPaneControl = new MainPanelEnityList();
            this._MainPanel = this.CustomTaskPanes.Add(taskPaneControl, "CRM Entities List");
            this._MainPanel.Width = 300;
            this._MainPanel.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
