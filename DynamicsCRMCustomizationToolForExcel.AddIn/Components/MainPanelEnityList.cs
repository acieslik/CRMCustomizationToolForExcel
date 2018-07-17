using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DynamicsCRMCustomizationToolForExcel.Controller;
using DynamicsCRMCustomizationToolForExcel.Model;
using System.Globalization;

namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    public partial class MainPanelEnityList : UserControl
    {
        private bool Isloading = false;
        private ComponentsTreeHandler treeHandler;
        public MainPanelEnityList()
        {
            InitializeComponent();
        }

        public void FillInstalledLangauges(int[] languages, int def)
        {
            cmbLanguage.Items.AddRange(languages.Select(x => ((new CultureInfo(x)).DisplayName)).ToArray());
            cmbLanguage.SelectedItem = (new CultureInfo(def)).DisplayName;
        }

        public void FillSolution(IEnumerable<Solution> sol, Solution def)
        {
            cmbSolution.Items.AddRange(sol.Select(x => x.SolutionName).ToArray());
            cmbSolution.SelectedItem = def.SolutionName;
        }


        public void FillEntitiesList()
        {
            Isloading = false;
            cmbSolution.Items.Clear();
            cmbLanguage.Items.Clear();
            GlobalApplicationData.Instance.enableSheetProtection = false;
            int language = GlobalApplicationData.Instance.currentLanguage;
            FillInstalledLangauges(GlobalApplicationData.Instance.crmInstalledLanguages, language);
            FillSolution(GlobalApplicationData.Instance.crmSolutions, GlobalApplicationData.Instance.currentSolution);
            treeHandler = new ComponentsTreeHandler(trvCrmComponents, GlobalApplicationData.Instance.currentEnitiesList,  GlobalOperations.Instance.CreatenNewFormSheet,  GlobalOperations.Instance.CreatenNewViewSheet,  GlobalOperations.Instance.CreatenNewAttributesSheet);
            treeHandler.addTreeViewContextMenuForViews(btnAddView, ctmShowView);
            UpdateEntityList(GlobalApplicationData.Instance.currentSolution.SolutionName);
            Isloading = false;
        }


        private void UpdateEntityList(string solution)
        {
            if (treeHandler != null)
            {
                 GlobalOperations.Instance.FilterBySolution(solution);
                treeHandler.refreshTree();
            }
        }

        private void cmbSolution_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSolution.SelectedItem.ToString() != string.Empty || !Isloading)
            {
                UpdateEntityList(cmbSolution.SelectedItem.ToString());
            }
        }

        private void btnOpenAll_Click(object sender, EventArgs e)
        {
            foreach (var item in GlobalApplicationData.Instance.currentEnitiesList)
            {
                if (item.MetadataId != null)
                {
                    GlobalOperations.Instance.CreatenNewAttributesSheet(item.MetadataId.Value);
                }
            }
        }

        private void cmbLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Isloading)
            {
                GlobalApplicationData.Instance.currentLanguage = GlobalApplicationData.Instance.crmInstalledLanguages[cmbLanguage.SelectedIndex];
                UpdateEntityList(GlobalApplicationData.Instance.currentSolution.SolutionName);
            }
        }

        private void chbSheetsProtection_CheckedChanged(object sender, EventArgs e)
        {
            GlobalApplicationData.Instance.enableSheetProtection = chbSheetsProtection.Checked;
        }

        private void tsmAddView_Click(object sender, EventArgs e)
        {
            //CrmViewDetails crmViewDetails = new CrmViewDetails();
            //crmViewDetails.ShowDialog();
        }

        private void trvCrmComponents_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void showViewAnfFetchXml_Click(object sender, EventArgs e)
        {
            //CrmViewFecthViewXml crmViewDetails = new CrmViewFecthViewXml();
            //crmViewDetails.ShowDialog();
        }

        private void pnlEntities_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }



    }

 
}
