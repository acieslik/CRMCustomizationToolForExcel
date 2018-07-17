using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using DynamicsCRMCustomizationToolForExcel.Model;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class OptionSetRequestGenerator
    {
        private int _languageCode;

        public int languageCode
        {
            get { return _languageCode; }
            set { _languageCode = value; }
        }

        private OptionSetExcelSheetsInfo optionMetadata;

        public OptionSetRequestGenerator(OptionSetExcelSheetsInfo optionMetadata)
        {
            this.languageCode = optionMetadata.language;
            this.optionMetadata = optionMetadata;
        }

        public IEnumerable<int> getDataMatrixValues(ExcelMatrix dataMatrix)
        {
            int optionValue;
            List<int> listValues = new List<int>();
            for (int i = 0; i < dataMatrix.numberofElements; i++)
            {
                if (int.TryParse(dataMatrix.getElement(i, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL), out optionValue))
                {
                    listValues.Add(optionValue);
                }
            }
            return listValues;
        }


        public IEnumerable<CrmOperation> generateCrmOperationRequest(ExcelMatrix dataMatrix)
        {
            int optionValue;
            bool changeOrder = false;
            int indexorder = 0;
            int[] valuesOrder = new int[dataMatrix.numberofElements];
            List<CrmOperation> crmOp = new List<CrmOperation>();
            IEnumerable<int> currentValues = getDataMatrixValues(dataMatrix);
            for (int i = 0; i < dataMatrix.numberofElements; i++)
            {
                if (int.TryParse(dataMatrix.getElement(i, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL), out optionValue))
                {
                    int nOptionExcel = currentValues.Where(x => x == optionValue).Count();
                    if (nOptionExcel > 1)
                    {
                        crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.error, CrmOperation.CrmOperationTarget.none, null, string.Format("Error , duplicate OptionSet Value:{0}", optionValue)));
                    }
                    else if (nOptionExcel == 1) 
                    { 
                        IEnumerable<OptionMetadata> option = optionMetadata.optionData.Options.Where(x => x.Value.Value == optionValue);
                        if (option.Count() == 0)
                        {
                            addOptionCreateRequest(dataMatrix.getRow(i), crmOp);
                        }
                        else 
                        {
                            addOptionUpdateRequest(dataMatrix.getRow(i), crmOp, option.First());
                        }

                        if (optionMetadata.optionData.Options.Count <= i || optionValue != optionMetadata.optionData.Options[i].Value)
                        {
                            changeOrder = true;
                        }
                        valuesOrder[indexorder++] = optionValue;
                    }

                }
                else
                {
                    crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.error, CrmOperation.CrmOperationTarget.none, null, string.Format("Error converting {0} to int", dataMatrix.getElement(i, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL))));
                }

            }
            checkOptionToRemove(currentValues, crmOp);
            if (changeOrder)
            {
                addOptionOrderRequestRequest(crmOp, valuesOrder);
            }
            //check revoved option
            return crmOp;
        }


        private void checkOptionToRemove(IEnumerable<int> optionValue, List<CrmOperation> crmOp)
        {
            foreach (var option in optionMetadata.optionData.Options)
            {
                if (option.Value != null && !optionValue.Contains(option.Value.Value))
                {
                    addOptionRemoveRequest(option.Value.Value, crmOp);
                }
            }
        }

        public void addOptionRemoveRequest(int value, List<CrmOperation> crmOp)
        {
            DeleteOptionValueRequest deleteOptionValueRequest;
            if (optionMetadata.optionData.IsGlobal.Value)
            {
                deleteOptionValueRequest =
                     new DeleteOptionValueRequest
                     {
                         OptionSetName = optionMetadata.optionData.Name,
                         Value = value
                     };
            }
            else
            {
                deleteOptionValueRequest =
                 new DeleteOptionValueRequest
                 {
                     AttributeLogicalName = optionMetadata.parentAttribute.LogicalName,
                     EntityLogicalName = optionMetadata.parentAttribute.EntityLogicalName,
                     Value = value
                 };
            }
            crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.optionSet, deleteOptionValueRequest, string.Format("Delete OptionSet {0}",value)));
        }


        public void addOptionCreateRequest(string[] row, List<CrmOperation> crmOp)
        {
            InsertOptionValueRequest insertOptionValueRequest;
            int result;
            if (optionMetadata.optionData.IsGlobal.Value)
            {
                insertOptionValueRequest =
                     new InsertOptionValueRequest
                        {
                            OptionSetName = optionMetadata.optionData.Name,
                            Label = new Label(row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL], languageCode),
                            Value = int.TryParse(row[ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL], out result) ? result : (int?)null
                        };
            }
            else
            {
                insertOptionValueRequest =
                 new InsertOptionValueRequest
                 {
                     AttributeLogicalName = optionMetadata.parentAttribute.LogicalName,
                     EntityLogicalName = optionMetadata.parentAttribute.EntityLogicalName,
                     Label = new Label(row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL], languageCode),
                     Value = int.TryParse(row[ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL], out result) ? result : (int?)null
                 };
            }
            crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.optionSet, insertOptionValueRequest, string.Format("Insert OptionSet {0}", row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL])));
        }

        public bool addOptionUpdateRequest(string[] row, List<CrmOperation> crmOp, OptionMetadata optionToCompare)
        {
            int valueresult;
            if (!int.TryParse(row[ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL], out valueresult))
            {
                return false;
            }
            if (Utils.getLocalizedLabel(optionToCompare.Label.LocalizedLabels,languageCode) != row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL]) 
            {
                UpdateOptionValueRequest updateOptionValueRequest;
                if (optionMetadata.optionData.IsGlobal.Value)
                {
                    updateOptionValueRequest =
                         new UpdateOptionValueRequest
                         {
                             OptionSetName = optionMetadata.optionData.Name,
                             Label = new Label(row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL], _languageCode),
                             MergeLabels = true,
                             Value = valueresult
                         };
                }
                else
                {
                    updateOptionValueRequest = new UpdateOptionValueRequest
                     {
                         AttributeLogicalName = optionMetadata.parentAttribute.LogicalName,
                         EntityLogicalName = optionMetadata.parentAttribute.EntityLogicalName,
                         Label = new Label(row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL], _languageCode),
                         MergeLabels = true,
                         Value = valueresult
                     };
                }
                crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.optionSet, updateOptionValueRequest, string.Format("Update OptionSet {0}", row[ExcelColumsDefinition.OPTIONSETLABELEXCELCOL])));
            }
            return true;
        }

        public void addOptionOrderRequestRequest(List<CrmOperation> crmOp, int[] optionValue)
        {
            OrderOptionRequest orderOptionRequest;
            if (optionMetadata.optionData.IsGlobal.Value)
            {
                orderOptionRequest =
                     new OrderOptionRequest
                     {
                         OptionSetName = optionMetadata.optionData.Name,
                         Values = optionValue
                     };
            }
            else
            {
                orderOptionRequest =
                 new OrderOptionRequest
                 {
                     AttributeLogicalName = optionMetadata.parentAttribute.LogicalName,
                     EntityLogicalName = optionMetadata.parentAttribute.EntityLogicalName,
                     Values = optionValue
                 };
            }
            crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.changeorder, CrmOperation.CrmOperationTarget.optionSet, orderOptionRequest, string.Format("Change Options order : {0}", string.Join(",", optionValue))));
        }

    }
}
