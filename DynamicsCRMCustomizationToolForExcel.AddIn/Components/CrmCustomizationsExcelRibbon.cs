using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk.Discovery;
using System.ServiceModel.Description;
using System.Net;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using DynamicsCRMCustomizationToolForExcel.Controller;
using DynamicsCRMCustomizationToolForExcel.Model;
using System.Threading.Tasks;

namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    public partial class CrmCustomizationsRibbon
    {

        public void enableRibbonButtons()
        {
            btnPublishAll.Enabled = true;
            btnUpdateSheet.Enabled = true;
            btnExportChanges.Enabled = true;
            btnSynchronizeEntity.Enabled = true;
            btnGetEntitiesList.Enabled = true;
        }


        private void btnConnect_Click(object sender, RibbonControlEventArgs e)
        {
             CrmLogin _ctrl = new CrmLogin();
            _ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
            _ctrl.ShowDialog();
        }

        private async void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLogin)
            {
               ((CrmLogin)sender).Close();
               Loading loading = new Loading();
               loading.Show();
               await Task.Run(() =>
               {
                   GlobalApplicationData.Instance.connectionInProgress = true;
                   OrganizationServiceProxy organizationProxy = ((CrmLogin)sender).CrmConnectionMgr.CrmSvc.OrganizationServiceProxy;
                   GlobalOperations.Instance.CRMOpHelper.Service = organizationProxy;
                   GlobalOperations.Instance.LoadOperations();
               });
               loading.Close();
               Globals.DynamicsCRMExcelAddIn.MainPanel.Visible = true;
               ((MainPanelEnityList)Globals.DynamicsCRMExcelAddIn.MainPanel.Control).FillEntitiesList();
               enableRibbonButtons();
               GlobalApplicationData.Instance.connectionInProgress = false;
            }
        }


        private void btnExportChanges_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalOperations.Instance.ExcelOperations.IsEditing())
            {
                MessageBox.Show("Please exit from edit-mode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ConfirmCustomization confirm = new ConfirmCustomization();
            confirm.ShowDialog();
        }


        private void btnUpdateSheet_Click(object sender, RibbonControlEventArgs e)
        {
            if ( GlobalOperations.Instance.ExcelOperations.IsEditing())
            {
                MessageBox.Show("Please exit from edit-mode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            GlobalOperations.Instance.RefreshCurrentSheet();
        }

      
        private void btnSynchronizeEntity_Click(object sender, RibbonControlEventArgs e)
        {
            //if (Globals.CrmAddIn.excelHelper.IsEditing())
            //{
            //    MessageBox.Show("Plesea exit from edit-mode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
            //Worksheet sheet = Globals.CrmAddIn.excelHelper.getCurrentSheet();
            //ExcelSheetInfo.ExcelSheetType type;
            //string name;
            //string orgprefix;
            //int language;
            //if (!Globals.CrmAddIn.excelHelper.readSettingRow(sheet, out name, out type, out orgprefix, out language)) return;
            //if (type == ExcelSheetInfo.ExcelSheetType.attribute)
            //{
            //    EntityMetadata currentEntity = Globals.CrmAddIn.crmOpHelper.RetriveEntityAtrribute(name);
            //    if (currentEntity != null)
            //    {
            //        GlobalApplicationData.Instance.eSheetsInfomation.addSheetAndSetAsCurrent(new AttributeExcelSheetsInfo(ExcelSheetInfo.ExcelSheetType.attribute, sheet, currentEntity), name);
            //        GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().orgPrefix = orgprefix;
            //        GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().language = language;
            //    }
            //    else
            //    {
            //        MessageBox.Show("Entity not Found Impossible To syncronize", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}
            //else if (type == ExcelSheetInfo.ExcelSheetType.optionSet)
            //{
            //    // to be implemented
            //}
        }

        private void btnGetEntitiesList_Click(object sender, RibbonControlEventArgs e)
        {
            //if (GlobalOperations.Instance.ExcelOperations.IsEditing())
            //{
            //    MessageBox.Show("Plesea exit from edit-mode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
            ////GlobalOperations.CreatenNewEntitySheet();
            //if (Globals.DynamicsCRMExcelAddIn.MainPanel.Visible == false)
            //{
            //    Globals.CrmAddIn.TaskPane.Visible = true;
            //    ((ActionPanelEntityList)Globals.CrmAddIn.TaskPane.Control).showLoading(true);
            //    GlobalOperations.LoadOperations();
            //    ((ActionPanelEntityList)Globals.CrmAddIn.TaskPane.Control).FillEntitiesList();
            //}
        }

        private async void btnPublishAll_Click(object sender, RibbonControlEventArgs e)
        {
            Loading loading = new Loading();
            loading.LabelText.Content = "Publishing...";
            loading.Show();

            Task work = ProcessAction(loading, GlobalOperations.Instance.CRMOpHelper.publishRequest);

            await work;
        }

        async Task ProcessAction(System.Windows.Window w, System.Action a)
        {
            await Task.Run(() => {
                a();
                Globals.DynamicsCRMExcelAddIn.SyncContext.Post(new System.Threading.SendOrPostCallback((o) => {
                    w.Close();
                }), null);
        });
            
        }
    }
}
