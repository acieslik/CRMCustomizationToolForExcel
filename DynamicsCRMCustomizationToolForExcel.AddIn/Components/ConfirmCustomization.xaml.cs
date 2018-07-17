using DynamicsCRMCustomizationToolForExcel.Controller;
using DynamicsCRMCustomizationToolForExcel.Model;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    /// <summary>
    /// Interaction logic for ConfirmCustomization.xaml
    /// </summary>
    public partial class ConfirmCustomization : Window
    {

        private ObservableCollection<CrmOperation> operationList;

        public ConfirmCustomization()
        {
            InitializeComponent();
        }

        public void btnExit_Click(object sender, RoutedEventArgs e)
        {
           MessageBoxResult res=  MessageBox.Show("Do you want to update the current sheet ?", "Update Sheet", MessageBoxButton.YesNoCancel);
           if (res == MessageBoxResult.No)
           {
               this.Close();
           }
           else
           {
               GlobalOperations.Instance.RefreshCurrentSheet();
               this.Close();
           }
        }

        public void frExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //prbCrmOperationStatus.Visible = false;
            ExcelSheetInfo currentsheet = GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet();
            string name;
            ExcelSheetInfo.ExcelSheetType type;
            string orgprefix;
            int language;
            if (GlobalOperations.Instance.ExcelOperations.readSettingRow(currentsheet.excelsheet, out name, out type, out orgprefix, out language))
            {
                currentsheet.orgPrefix = orgprefix;
                currentsheet.language = language;
            }

            if (currentsheet != null)
            {
                operationList = new ObservableCollection<CrmOperation>(generateOperationCurrentSheet(currentsheet));
                if (operationList != null)
                {
                    ShowGridData(false);
                }
            }
        }



        private IEnumerable<CrmOperation> generateOperationCurrentSheet(ExcelSheetInfo currentSheet)
        {
            ExcelMatrix matrix;
            switch (currentSheet.sheetType)
            {
                case ExcelSheetInfo.ExcelSheetType.attribute:
                    matrix = GlobalOperations.Instance.ExcelOperations.getExcelDataMatrix(currentSheet.excelsheet, ExcelColumsDefinition.MAXNUMBEROFCOLUMN, ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE, ExcelColumsDefinition.SCHEMANAMEEXCELCOL);
                    AttributeRequestGenerator reqGeneratorHelper = new AttributeRequestGenerator((AttributeExcelSheetsInfo)currentSheet, GlobalApplicationData.Instance.optionSetData);
                    return reqGeneratorHelper.generateCrmOperationRequest(matrix);
                case ExcelSheetInfo.ExcelSheetType.optionSet:
                    matrix = GlobalOperations.Instance.ExcelOperations.getExcelDataMatrix(currentSheet.excelsheet, ExcelColumsDefinition.MAXNUMBEROFCOLUMN, ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL);
                    OptionSetRequestGenerator optGenerator = new OptionSetRequestGenerator((OptionSetExcelSheetsInfo)currentSheet);
                    return optGenerator.generateCrmOperationRequest(matrix);
                case ExcelSheetInfo.ExcelSheetType.entity:
                    matrix = GlobalOperations.Instance.ExcelOperations.getExcelDataMatrix(currentSheet.excelsheet, ExcelColumsDefinition.MAXNUMBEROFCOLUMN, ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE, ExcelColumsDefinition.ENTITYSCHEMANAMEEXCELCOL);
                    EntityRequestGenerator entGenerator = new EntityRequestGenerator((EntityExcelSheetsInfo)currentSheet);
                    return entGenerator.generateCrmOperationRequest(matrix);
                case ExcelSheetInfo.ExcelSheetType.view:
                    MessageBox.Show("The view editor feature is released as Alpha Version. Please ensure to back up your CRM solution and report any issue on Codeplex project page.");
                    matrix = GlobalOperations.Instance.ExcelOperations.getExcelDataMatrix(currentSheet.excelsheet, ExcelColumsDefinition.MAXNUMBEROFCOLUMN, ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE, ExcelColumsDefinition.VIEWATTRIBUTENAME);
                    return ViewXml.generateCrmOperationRequest(matrix, (ViewExcelSheetsInfo)currentSheet);
                case ExcelSheetInfo.ExcelSheetType.form:
                    MessageBox.Show("Forms are read only");
                    this.Close();
                    return new List<CrmOperation>();
            }
            return null;
        }
            
        private void ShowGridData(bool showResult)
        {
            if (showResult)
            {
                GridView gridview = lstOperationsList.View as GridView;
                gridview.Columns[3].Width = 125;
                gridview.Columns[4].Width = 75;
                gridview.Columns[5].Width = 75;
            }
            lstOperationsList.ItemsSource = operationList;
            lstOperationsList.Items.Refresh();
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += WorkerDoWork;
            worker.ProgressChanged += WorkerReportProgress;
            worker.RunWorkerCompleted += WorkerReporCompleted;
            worker.RunWorkerAsync();
            prbCrmOperationStatus.Visibility = Visibility.Visible;

        }

        private void WorkerDoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            int i = 0;
            foreach (var item in operationList)
            {
                if (item.executeOperation)
                {
                    GlobalOperations.Instance.CRMOpHelper.executeOpertionsCrm(item);
                    i++;
                }
                worker.ReportProgress((int)(((double)i / operationList.Count()) * 100));
            }

            ExcelSheetInfo currentsheet = GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet();
            if (currentsheet != null)
            {
                switch (currentsheet.sheetType)
                {
                    case ExcelSheetInfo.ExcelSheetType.view:
                        {
                            ViewExcelSheetsInfo viewSheet = (ViewExcelSheetsInfo)currentsheet;
                            if (viewSheet.isNew == true && operationList.First().operationSucceded)
                            {
                                if (operationList.Count() > 0 && operationList.First().orgResponse != null && operationList.First().orgResponse is CreateResponse)
                                {
                                    viewSheet.viewId = ((CreateResponse)operationList.First().orgResponse).id;
                                }
                                viewSheet.isNew = false;
                            }
                            break;
                        }
                }
            }

        }
        private void WorkerReportProgress(object sender, ProgressChangedEventArgs e)
        {
            prbCrmOperationStatus.Value = e.ProgressPercentage;
        }

        private void WorkerReporCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ShowGridData(true);
        }
    }
}
