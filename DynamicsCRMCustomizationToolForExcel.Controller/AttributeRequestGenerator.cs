using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using DynamicsCRMCustomizationToolForExcel.Model;
using System.Runtime.Serialization;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class AttributeRequestGenerator
    {
        private AttributeMetadata[] _filteredMetadata;

        public AttributeMetadata[] filteredMetadata
        {
            get { return _filteredMetadata; }
            set { _filteredMetadata = value; }
        }

        private string _entityLocialName;

        public string entityLocialName
        {
            get { return _entityLocialName; }
            set { _entityLocialName = value; }
        }

        private bool currentOperationCreate;

        private string organizationPrefix;
        private int languageCode;

        private OptionSetMetadataBase [] _optionSetData;

        public OptionSetMetadataBase [] optionSetData
        {
            get { return _optionSetData; }
            set { _optionSetData = value; }
        }


        public AttributeRequestGenerator(AttributeExcelSheetsInfo sheet, OptionSetMetadataBase [] optionSetData)
        {
            languageCode = sheet.language;
            organizationPrefix = sheet.orgPrefix;
            this.filteredMetadata = sheet.entityMedata.Attributes;
            this.entityLocialName = sheet.entityMedata.LogicalName;
            this.optionSetData = optionSetData;
        }

        #region  field creation

        private AttributeMetadata generalFieldCreation(string[] row, AttributeMetadata attrMetadata)
        {
            AttributeRequiredLevel attributeRequiredLevel = AttributeRequiredLevel.None;

            String reqLevelstring = row[ExcelColumsDefinition.REQUIREDLEVELEXCELCOL];
            if (Enum.IsDefined(typeof(AttributeRequiredLevel), reqLevelstring))
                attributeRequiredLevel = (AttributeRequiredLevel)Enum.Parse(typeof(AttributeRequiredLevel), reqLevelstring, true);
            else
                attributeRequiredLevel = AttributeRequiredLevel.None;
            bool parseresult;
            attrMetadata.SchemaName = Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate);
            attrMetadata.DisplayName = new Microsoft.Xrm.Sdk.Label(row[ExcelColumsDefinition.DISPLAYNAMEEXCELCOL], languageCode);
            attrMetadata.Description = new Microsoft.Xrm.Sdk.Label(row[ExcelColumsDefinition.DESCRIPTIONEXCELCOL], languageCode);
            attrMetadata.RequiredLevel = new AttributeRequiredLevelManagedProperty(attributeRequiredLevel);
            attrMetadata.IsValidForAdvancedFind = Boolean.TryParse(row[ExcelColumsDefinition.ADVANCEDFINF], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(true);
            attrMetadata.IsSecured = Boolean.TryParse(row[ExcelColumsDefinition.SECURED], out parseresult) ? parseresult : true;
            attrMetadata.IsAuditEnabled = Boolean.TryParse(row[ExcelColumsDefinition.AUDITENABLED], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(false);
            return attrMetadata;
        }


        private AttributeMetadata stringFieldCreation(string[] row)
        {
            StringAttributeMetadata attrMetadata = new StringAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            int parseResult;
            attrMetadata.MaxLength = int.TryParse(row[ExcelColumsDefinition.STRINGMAXLENGTHCOL], out parseResult) ? parseResult : 100;
            string imeModeFormatString = row[ExcelColumsDefinition.STRINGIMEMODECOL];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeFormatString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeFormatString, true) : (ImeMode?)null;
            string formatString = row[ExcelColumsDefinition.STRINGFORMATCOL];
            StringFormat? stringFormat = Enum.IsDefined(typeof(StringFormat), formatString) ? (StringFormat)Enum.Parse(typeof(StringFormat), formatString, true) : StringFormat.Text;
            attrMetadata.Format = stringFormat;
            return attrMetadata;
        }

        private AttributeMetadata integerFieldCreation(string[] row)
        {
            IntegerAttributeMetadata attrMetadata = new IntegerAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            int parseResult;
            string intgerFormatString = row[ExcelColumsDefinition.INTEGERFORMATCOL];
            attrMetadata.Format = Enum.IsDefined(typeof(IntegerFormat), intgerFormatString) ? (IntegerFormat)Enum.Parse(typeof(IntegerFormat), intgerFormatString, true) : (IntegerFormat?)null;
            attrMetadata.MaxValue = int.TryParse(row[ExcelColumsDefinition.INTEGERMAXVALUECOL], out parseResult) ? parseResult : (int?)null;
            attrMetadata.MinValue = int.TryParse(row[ExcelColumsDefinition.INTEGERMINVALUECOL], out parseResult) ? parseResult : (int?)null;
            return attrMetadata;
        }

        private AttributeMetadata dateTimeFieldCreation(string[] row)
        {
            string datetimeFormatString = row[ExcelColumsDefinition.DATETIMEFORMATCOL];
            Microsoft.Xrm.Sdk.Metadata.DateTimeFormat? dateTimeFormat = Enum.IsDefined(typeof(Microsoft.Xrm.Sdk.Metadata.DateTimeFormat), datetimeFormatString) ? (Microsoft.Xrm.Sdk.Metadata.DateTimeFormat)Enum.Parse(typeof(Microsoft.Xrm.Sdk.Metadata.DateTimeFormat), datetimeFormatString, true) : (Microsoft.Xrm.Sdk.Metadata.DateTimeFormat?)null;
            DateTimeAttributeMetadata attrMetadata = new DateTimeAttributeMetadata(dateTimeFormat, Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            string imeModeFormatString = row[ExcelColumsDefinition.DATETIMEIMEMODECOL];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeFormatString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeFormatString, true) : (ImeMode?)null;
            return attrMetadata;
        }

        private AttributeMetadata doubleFieldCreation(string[] row)
        {
            DoubleAttributeMetadata attrMetadata = new DoubleAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            double doubleResult;
            int parseResult;
            attrMetadata.Precision = int.TryParse(row[ExcelColumsDefinition.DOUBLEPRECISION], out parseResult) ? parseResult : (int?)null;
            attrMetadata.MaxValue = double.TryParse(row[ExcelColumsDefinition.DOUBLEMAXVALUE], out doubleResult) ? doubleResult : (double?)null;
            attrMetadata.MinValue = double.TryParse(row[ExcelColumsDefinition.DOUBLEMINVALUE], out doubleResult) ? doubleResult : (double?)null;
            string imeModeFormatString = row[ExcelColumsDefinition.DOUBLEIMEMODE];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeFormatString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeFormatString, true) : (ImeMode?)null;
            return attrMetadata;
        }

        private AttributeMetadata memoFieldCreation(string[] row)
        {
            MemoAttributeMetadata attrMetadata = new MemoAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            string memoFormatString = row[ExcelColumsDefinition.MEMOFORMATCOL];
            int parseResult;
            attrMetadata.Format = Enum.IsDefined(typeof(StringFormat), memoFormatString) ? (StringFormat)Enum.Parse(typeof(StringFormat), memoFormatString, true) : (StringFormat?)null;
            string imeModeString = row[ExcelColumsDefinition.MEMOIMEMODECOL];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeString, true) : (ImeMode?)null;
            attrMetadata.MaxLength = int.TryParse(row[ExcelColumsDefinition.MEMOMAXLENGTHCOL], out parseResult) ? parseResult : 2000;
            return attrMetadata;
        }

        private AttributeMetadata pickListFieldCreation(string[] row)
        {
            PicklistAttributeMetadata attrMetadata = new PicklistAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            int parseResult;
            attrMetadata.DefaultFormValue = int.TryParse(row[ExcelColumsDefinition.PICKLISTDEFAULTVALUE], out parseResult) ? parseResult : (int?)null;
            bool globalOption;
            globalOption = Boolean.TryParse(row[ExcelColumsDefinition.PICKLISTGLOBAL], out globalOption) ? globalOption : false;
            if (globalOption)
            {
                attrMetadata.OptionSet = new OptionSetMetadata();
                attrMetadata.OptionSet.IsGlobal = true;
                IEnumerable<OptionSetMetadataBase> selectdeoption = optionSetData.Where(x => x.Name == row[ExcelColumsDefinition.PICKLISTGLOBALNAME]);
                attrMetadata.OptionSet.Name = selectdeoption.Count() > 0 ? selectdeoption.First().Name : null;
            }
            else
            {
                attrMetadata.OptionSet = new OptionSetMetadata();
                attrMetadata.OptionSet.IsGlobal = false;
            }
            return attrMetadata;
        }

        private AttributeMetadata decimalFieldCreation(string[] row)
        {
            DecimalAttributeMetadata attrMetadata = new DecimalAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            int parseResult;
            decimal decimalResult;
            attrMetadata.MaxValue = decimal.TryParse(row[ExcelColumsDefinition.DECIMALMAXVALUE], out decimalResult) ? decimalResult : (decimal?)null;
            attrMetadata.MinValue = decimal.TryParse(row[ExcelColumsDefinition.DECIMALMINVALUE], out decimalResult) ? decimalResult : (decimal?)null;
            attrMetadata.Precision = int.TryParse(row[ExcelColumsDefinition.DECIMALPRECISION], out parseResult) ? parseResult : (int?)null;
            string imeModeFormatString = row[ExcelColumsDefinition.DECIMALIMEMODE];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeFormatString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeFormatString, true) : (ImeMode?)null;
            return attrMetadata;
        }

        private AttributeMetadata booleanFieldCreation(string[] row)
        {
            bool parseresult;
            BooleanAttributeMetadata attrMetadata = new BooleanAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            attrMetadata.DefaultValue = Boolean.TryParse(row[ExcelColumsDefinition.BOOLEANDEFAULTVALUE], out parseresult) ? parseresult : false;
            attrMetadata.OptionSet = new BooleanOptionSetMetadata();
            attrMetadata.OptionSet.TrueOption = row[ExcelColumsDefinition.BOOLEANTRUEOPTION] != string.Empty ? new OptionMetadata(new Label(row[ExcelColumsDefinition.BOOLEANTRUEOPTION], languageCode), 1) : new OptionMetadata(new Label(string.Empty, languageCode), 1);
            attrMetadata.OptionSet.FalseOption = row[ExcelColumsDefinition.BOOLEANFALSEOPTION] != string.Empty ? new OptionMetadata(new Label(row[ExcelColumsDefinition.BOOLEANFALSEOPTION], languageCode), 0) : new OptionMetadata(new Label(string.Empty, languageCode), 0);
            return attrMetadata;
        }

        private AttributeMetadata moneyFieldCreation(string[] row)
        {
            MoneyAttributeMetadata attrMetadata = new MoneyAttributeMetadata(Utils.addOrgPrefix(row[ExcelColumsDefinition.SCHEMANAMEEXCELCOL], organizationPrefix, currentOperationCreate));
            generalFieldCreation(row, attrMetadata);
            int parseResult;
            double doubleResult;
            attrMetadata.Precision = int.TryParse(row[ExcelColumsDefinition.MONEYPRECISION], out parseResult) ? parseResult : (int?)null;
            attrMetadata.MaxValue = double.TryParse(row[ExcelColumsDefinition.MONEYMAXVALUE], out doubleResult) ? doubleResult : (double?)null;
            attrMetadata.MinValue = double.TryParse(row[ExcelColumsDefinition.MONEYMINVALUE], out doubleResult) ? doubleResult : (double?)null;
            string imeModeFormatString = row[ExcelColumsDefinition.MONEYIMEMODE];
            attrMetadata.ImeMode = Enum.IsDefined(typeof(ImeMode), imeModeFormatString) ? (ImeMode)Enum.Parse(typeof(ImeMode), imeModeFormatString, true) : (ImeMode?)null;
            return attrMetadata;
        }

        private IExtensibleDataObject lookUpFieldCreation(string[] row)
        {
            LookupAttributeMetadata attrMetadata = new LookupAttributeMetadata();
            generalFieldCreation(row, attrMetadata);
            OneToManyRelationshipMetadata oneToManyRelationship = new OneToManyRelationshipMetadata();
            oneToManyRelationship.ReferencingEntity = entityLocialName;
            string relatiosshipName = row[ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME] != string.Empty ? row[ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME] : string.Empty;
            string relashionshiptarget = row[ExcelColumsDefinition.LOOKUPTARGET] != string.Empty ? row[ExcelColumsDefinition.LOOKUPTARGET] : string.Empty;
            oneToManyRelationship.ReferencedEntity = relashionshiptarget;
            oneToManyRelationship.SchemaName = Utils.addOrgPrefix(relatiosshipName, organizationPrefix, currentOperationCreate);
            CreateOneToManyRequest createRelationship = new CreateOneToManyRequest();
            createRelationship.Lookup = attrMetadata;
            createRelationship.OneToManyRelationship = oneToManyRelationship;
            return createRelationship;
        }




        private IEnumerable<IExtensibleDataObject> pickListOptionCreation(string[] row, PicklistAttributeMetadata attr)
        {
            List<IExtensibleDataObject> opList = new List<IExtensibleDataObject>();
            int result;
            string optionString = row[ExcelColumsDefinition.PICKLISTREF];
            if (optionString != string.Empty && optionString != ExcelDataAccess.voidOptionSetString)
            {
                foreach (string optionLine in optionString.Split(ExcelDataAccess.optionSetSeparator))
                {
                    InsertOptionValueRequest insertOptionValueRequest = new InsertOptionValueRequest();
                    string[] optValueName = optionLine.Split(new Char[1] { ExcelDataAccess.optionSetNumberSeparator }, 2);
                    if (optValueName.Count() == 2)
                    {
                        if (attr.OptionSet.IsGlobal != null && attr.OptionSet.IsGlobal.Value)
                        {
                            insertOptionValueRequest =
                                 new InsertOptionValueRequest
                                 {
                                     OptionSetName = attr.OptionSet.Name,
                                     Label = new Label(optValueName[1], languageCode),
                                     Value = int.TryParse(optValueName[0], out result) ? result : (int?)null
                                 };
                        }
                        else
                        {
                            insertOptionValueRequest =
                             new InsertOptionValueRequest
                             {
                                 AttributeLogicalName = attr.SchemaName.ToLower(),
                                 EntityLogicalName = entityLocialName,
                                 Label = new Label(optValueName[1], languageCode),
                                 Value = int.TryParse(optValueName[0], out result) ? result : (int?)null
                             };
                        }
                        opList.Add(insertOptionValueRequest);
                    }
                }
            }
            return opList;
        }
        #endregion

        private IEnumerable<IExtensibleDataObject> attributeReader(string[] row)
        {
            List<IExtensibleDataObject> opList = new List<IExtensibleDataObject>();
            IExtensibleDataObject attrMetadata;
            switch (row[ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL])
            {
                case "Integer":
                    attrMetadata = integerFieldCreation(row);
                    break;
                case "Boolean":
                    attrMetadata = booleanFieldCreation(row);
                    break;
                case "DateTime":
                    attrMetadata = dateTimeFieldCreation(row);
                    break;
                case "String":
                    attrMetadata = stringFieldCreation(row);
                    break;
                case "Picklist":
                    attrMetadata = pickListFieldCreation(row);
                    if (attrMetadata != null)
                    {
                        opList.AddRange(pickListOptionCreation(row, (PicklistAttributeMetadata)attrMetadata));
                    }
                    break;
                case "Memo":
                    attrMetadata = memoFieldCreation(row);
                    break;
                case "Double":
                    attrMetadata = doubleFieldCreation(row);
                    break;
                case "Decimal":
                    attrMetadata = decimalFieldCreation(row);
                    break;
                case "Money":
                    attrMetadata = moneyFieldCreation(row);
                    break;
                case "Lookup":
                    attrMetadata = lookUpFieldCreation(row);
                    break;
                default:
                    AttributeMetadata atMetadata = new AttributeMetadata();
                    attrMetadata = generalFieldCreation(row, atMetadata);
                    break;
            }
            opList.Insert(0, attrMetadata);
            return opList;
        }

        private void addCreateRequest(string[] row, List<CrmOperation> crmOp)
        {
            IEnumerable<IExtensibleDataObject> attrMetadataList = attributeReader(row);
            foreach (var attrMetadata in attrMetadataList)
            {
                if (attrMetadata != null)
                {
                    if (attrMetadata is AttributeMetadata)
                    {
                        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                        {
                            EntityName = filteredMetadata[0].EntityLogicalName,
                            Attribute = attrMetadata as AttributeMetadata,
                            SolutionUniqueName = GlobalApplicationData.Instance.currentSolution.SolutionName
                        };
                        crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.attribute, createAttributeRequest, "Create field " + ((AttributeMetadata)attrMetadata).SchemaName + " in " + filteredMetadata[0].EntityLogicalName));
                    }
                    else if (attrMetadata is CreateOneToManyRequest)
                    {
                        string outputstring = string.Format("Create relation {0} -> {1}", ((CreateOneToManyRequest)attrMetadata).OneToManyRelationship.ReferencingEntity, ((CreateOneToManyRequest)attrMetadata).OneToManyRelationship.ReferencedEntity);
                        crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.attribute, attrMetadata as CreateOneToManyRequest, outputstring));
                    }
                    else if (attrMetadata is InsertOptionValueRequest)
                    {
                        string outputstring = string.Format("Insert Option to {0}  Text:  {1}", ((InsertOptionValueRequest)attrMetadata).OptionSetName, Utils.getLocalizedLabel(((InsertOptionValueRequest)attrMetadata).Label.LocalizedLabels, languageCode));
                        crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.attribute, attrMetadata as InsertOptionValueRequest, outputstring));
                    }
                }
            }
        }


        private void addUpdateRequest(string[] row, List<CrmOperation> crmOp, AttributeMetadata currentAttribute)
        {
            IEnumerable<IExtensibleDataObject> excelAttrMetadataList = attributeReader(row);
            if (excelAttrMetadataList.Count() > 0)
            {
                IExtensibleDataObject excelAttrMetadata = excelAttrMetadataList.First();
                if (excelAttrMetadata != null)
                {
                    if (excelAttrMetadata is AttributeMetadata || excelAttrMetadata is CreateOneToManyRequest)
                    {
                        if (excelAttrMetadata is CreateOneToManyRequest)
                        {
                            excelAttrMetadata = ((CreateOneToManyRequest)excelAttrMetadata).Lookup;
                        }
                        IEnumerable<string> fieldToCahnge = checkDifferenceAttributeMetadata(currentAttribute, excelAttrMetadata as AttributeMetadata);
                        if (fieldToCahnge != null && fieldToCahnge.Count() != 0)
                        {
                            UpdateAttributeRequest updateAttributeRequest = new UpdateAttributeRequest
                            {
                                EntityName = filteredMetadata[0].EntityLogicalName,
                                Attribute = currentAttribute,
                                MergeLabels = false
                            };
                            crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.update, CrmOperation.CrmOperationTarget.attribute, updateAttributeRequest, "Update field " + currentAttribute.SchemaName + " in " + filteredMetadata[0].EntityLogicalName));
                        }
                    }
                }
            }
        }


        public IEnumerable<CrmOperation> generateCrmOperationRequest(ExcelMatrix dataMatrix)
        {
            List<CrmOperation> crmOp = new List<CrmOperation>();
            for (int i = 0; i < dataMatrix.numberofElements; i++)
            {
                IEnumerable<AttributeMetadata> attribute = filteredMetadata.Where(x => x.SchemaName == dataMatrix.getElement(i, ExcelColumsDefinition.SCHEMANAMEEXCELCOL));
                if (attribute.Count() == 0)
                {
                    currentOperationCreate = true;
                    addCreateRequest(dataMatrix.getRow(i), crmOp);
                }
                else if (attribute.Count() == 1)
                {
                    currentOperationCreate = false;
                    addUpdateRequest(dataMatrix.getRow(i), crmOp, attribute.First());
                }
            }
            return crmOp;
        }

        #region update attributes
        private IEnumerable<string> checkDifferenceAttributeMetadata(AttributeMetadata originalAttributeMetadata, AttributeMetadata readAttributeMetadata)
        {
            List<string> attributesToUpdate = checkGlobalDifferenceAttributeMetadata(originalAttributeMetadata, readAttributeMetadata);
            switch (originalAttributeMetadata.AttributeType)
            {
                case AttributeTypeCode.Integer:
                    IntegerAttributeMetadata intattrMetadata = originalAttributeMetadata as IntegerAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceIntegerAttribute(intattrMetadata, readAttributeMetadata as IntegerAttributeMetadata));
                    originalAttributeMetadata = intattrMetadata;
                    break;
                case AttributeTypeCode.DateTime:
                    DateTimeAttributeMetadata dateattrMetadata = originalAttributeMetadata as DateTimeAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceDateTimeAttribute(dateattrMetadata, readAttributeMetadata as DateTimeAttributeMetadata));
                    originalAttributeMetadata = dateattrMetadata;
                    break;
                case AttributeTypeCode.String:
                    StringAttributeMetadata strattrMetadata = originalAttributeMetadata as StringAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceStringAttribute(strattrMetadata, readAttributeMetadata as StringAttributeMetadata));
                    originalAttributeMetadata = strattrMetadata;
                    break;
                case AttributeTypeCode.Picklist:
                    PicklistAttributeMetadata pklattrMetadata = originalAttributeMetadata as PicklistAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferencePicklistAttribute(pklattrMetadata, readAttributeMetadata as PicklistAttributeMetadata));
                    originalAttributeMetadata = pklattrMetadata;
                    break;
                case AttributeTypeCode.Memo:
                    MemoAttributeMetadata memattrMetadata = originalAttributeMetadata as MemoAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceMemoAttribute(memattrMetadata, readAttributeMetadata as MemoAttributeMetadata));
                    originalAttributeMetadata = memattrMetadata;
                    break;
                case AttributeTypeCode.Double:
                    DoubleAttributeMetadata dblattrMetadata = originalAttributeMetadata as DoubleAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceDoubleAttribute(dblattrMetadata, readAttributeMetadata as DoubleAttributeMetadata));
                    originalAttributeMetadata = dblattrMetadata;
                    break;
                case AttributeTypeCode.Decimal:
                    DecimalAttributeMetadata dcmattrMetadata = originalAttributeMetadata as DecimalAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceDecimalAttribute(dcmattrMetadata, readAttributeMetadata as DecimalAttributeMetadata));
                    originalAttributeMetadata = dcmattrMetadata;
                    break;
                case AttributeTypeCode.Boolean:
                    BooleanAttributeMetadata blnattrMetadata = originalAttributeMetadata as BooleanAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceBooleanAttribute(blnattrMetadata, readAttributeMetadata as BooleanAttributeMetadata));
                    originalAttributeMetadata = blnattrMetadata;
                    break;
                case AttributeTypeCode.Money:
                    MoneyAttributeMetadata mnyattrMetadata = originalAttributeMetadata as MoneyAttributeMetadata;
                    attributesToUpdate.AddRange(checkDifferenceMoneyAttribute(mnyattrMetadata, readAttributeMetadata as MoneyAttributeMetadata));
                    originalAttributeMetadata = mnyattrMetadata;
                    break;
            }
            return attributesToUpdate;
        }

        private IEnumerable<string> checkDifferenceIntegerAttribute(IntegerAttributeMetadata originalAttributeMetadata, IntegerAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.MinValue != readAttributeMetadata.MinValue)
            {
                originalAttributeMetadata.MinValue = readAttributeMetadata.MinValue;
                attributeToChange.Add("MinValue");
            }
            if (originalAttributeMetadata.MaxValue != readAttributeMetadata.MaxValue)
            {
                originalAttributeMetadata.MaxValue = readAttributeMetadata.MaxValue;
                attributeToChange.Add("MaxValue");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceDateTimeAttribute(DateTimeAttributeMetadata originalAttributeMetadata, DateTimeAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.Format != readAttributeMetadata.Format)
            {
                originalAttributeMetadata.Format = readAttributeMetadata.Format;
                attributeToChange.Add("Format");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceStringAttribute(StringAttributeMetadata originalAttributeMetadata, StringAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.MaxLength != readAttributeMetadata.MaxLength)
            {
                originalAttributeMetadata.MaxLength = readAttributeMetadata.MaxLength;
                attributeToChange.Add("MaxLength");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;


        }

        private IEnumerable<string> checkDifferencePicklistAttribute(PicklistAttributeMetadata originalAttributeMetadata, PicklistAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.DefaultFormValue != readAttributeMetadata.DefaultFormValue)
            {
                originalAttributeMetadata.DefaultFormValue = readAttributeMetadata.DefaultFormValue;
                attributeToChange.Add("DefaultFormValue");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceMemoAttribute(MemoAttributeMetadata originalAttributeMetadata, MemoAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.MaxLength != readAttributeMetadata.MaxLength)
            {
                originalAttributeMetadata.MaxLength = readAttributeMetadata.MaxLength;
                attributeToChange.Add("MaxLength");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceDoubleAttribute(DoubleAttributeMetadata originalAttributeMetadata, DoubleAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.Precision != readAttributeMetadata.Precision)
            {
                originalAttributeMetadata.Precision = readAttributeMetadata.Precision;
                attributeToChange.Add("Precision");
            }
            if (originalAttributeMetadata.MinValue != readAttributeMetadata.MinValue)
            {
                originalAttributeMetadata.MinValue = readAttributeMetadata.MinValue;
                attributeToChange.Add("Min Value");
            }
            if (originalAttributeMetadata.MaxValue != readAttributeMetadata.MaxValue)
            {
                originalAttributeMetadata.MaxValue = readAttributeMetadata.MaxValue;
                attributeToChange.Add("Max Value");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceDecimalAttribute(DecimalAttributeMetadata originalAttributeMetadata, DecimalAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.Precision != readAttributeMetadata.Precision)
            {
                originalAttributeMetadata.Precision = readAttributeMetadata.Precision;
                attributeToChange.Add("Precision");
            }
            if (originalAttributeMetadata.MinValue != readAttributeMetadata.MinValue)
            {
                originalAttributeMetadata.MinValue = readAttributeMetadata.MinValue;
                attributeToChange.Add("Min Value");
            }
            if (originalAttributeMetadata.MaxValue != readAttributeMetadata.MaxValue)
            {
                originalAttributeMetadata.MaxValue = readAttributeMetadata.MaxValue;
                attributeToChange.Add("Max Value");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceBooleanAttribute(BooleanAttributeMetadata originalAttributeMetadata, BooleanAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.DefaultValue != readAttributeMetadata.DefaultValue)
            {
                originalAttributeMetadata.DefaultValue = readAttributeMetadata.DefaultValue;
                attributeToChange.Add("DefaultValue");
            }
            if (Utils.getLocalizedLabel(originalAttributeMetadata.OptionSet.TrueOption.Label.LocalizedLabels, languageCode) != Utils.getLocalizedLabel(readAttributeMetadata.OptionSet.TrueOption.Label.LocalizedLabels, languageCode))
            {
                Utils.setLocalizedLabel(originalAttributeMetadata.OptionSet.TrueOption.Label.LocalizedLabels, languageCode, Utils.getLocalizedLabel(readAttributeMetadata.OptionSet.TrueOption.Label.LocalizedLabels, languageCode));
                attributeToChange.Add("True Option");
            }
            if (Utils.getLocalizedLabel(originalAttributeMetadata.OptionSet.FalseOption.Label.LocalizedLabels, languageCode) != Utils.getLocalizedLabel(readAttributeMetadata.OptionSet.FalseOption.Label.LocalizedLabels, languageCode))
            {
                Utils.setLocalizedLabel(originalAttributeMetadata.OptionSet.FalseOption.Label.LocalizedLabels, languageCode, Utils.getLocalizedLabel(readAttributeMetadata.OptionSet.FalseOption.Label.LocalizedLabels, languageCode));
                attributeToChange.Add("False Option");
            }
            return attributeToChange;
        }

        private IEnumerable<string> checkDifferenceMoneyAttribute(MoneyAttributeMetadata originalAttributeMetadata, MoneyAttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();
            if (originalAttributeMetadata.Precision != readAttributeMetadata.Precision)
            {
                originalAttributeMetadata.Precision = readAttributeMetadata.Precision;
                attributeToChange.Add("Precision");
            }
            if (originalAttributeMetadata.MinValue != readAttributeMetadata.MinValue)
            {
                originalAttributeMetadata.MinValue = readAttributeMetadata.MinValue;
                attributeToChange.Add("Min Value");
            }
            if (originalAttributeMetadata.MaxValue != readAttributeMetadata.MaxValue)
            {
                originalAttributeMetadata.MaxValue = readAttributeMetadata.MaxValue;
                attributeToChange.Add("Max Value");
            }
            if (originalAttributeMetadata.ImeMode != readAttributeMetadata.ImeMode)
            {
                originalAttributeMetadata.ImeMode = readAttributeMetadata.ImeMode;
                attributeToChange.Add("ImeMode");
            }
            return attributeToChange;
        }

      
        private List<string> checkGlobalDifferenceAttributeMetadata(AttributeMetadata originalAttributeMetadata, AttributeMetadata readAttributeMetadata)
        {
            List<string> attributeToChange = new List<string>();

            if (Utils.getLocalizedLabel(originalAttributeMetadata.DisplayName.LocalizedLabels, languageCode) != Utils.getLocalizedLabel(readAttributeMetadata.DisplayName.LocalizedLabels, languageCode))
            {
                Utils.setLocalizedLabel(originalAttributeMetadata.DisplayName.LocalizedLabels, languageCode, Utils.getLocalizedLabel(readAttributeMetadata.DisplayName.LocalizedLabels, languageCode));
                attributeToChange.Add("Display Name");
            }
            if (Utils.getLocalizedLabel(originalAttributeMetadata.Description.LocalizedLabels, languageCode) != Utils.getLocalizedLabel(readAttributeMetadata.Description.LocalizedLabels, languageCode)) 
            {
                Utils.setLocalizedLabel(originalAttributeMetadata.Description.LocalizedLabels, languageCode, Utils.getLocalizedLabel(readAttributeMetadata.Description.LocalizedLabels, languageCode));
                attributeToChange.Add("Description");
            }

            if (originalAttributeMetadata.RequiredLevel.Value != readAttributeMetadata.RequiredLevel.Value)
            {
                originalAttributeMetadata.RequiredLevel = readAttributeMetadata.RequiredLevel;
                attributeToChange.Add("RequiredLevel");
            }
            if (originalAttributeMetadata.IsValidForAdvancedFind.Value != readAttributeMetadata.IsValidForAdvancedFind.Value)
            {
                originalAttributeMetadata.IsValidForAdvancedFind = readAttributeMetadata.IsValidForAdvancedFind;
                attributeToChange.Add("IsValidForAdvancedFind");
            }
            if (originalAttributeMetadata.IsSecured.Value != readAttributeMetadata.IsSecured.Value)
            {
                originalAttributeMetadata.IsSecured = readAttributeMetadata.IsSecured;
                attributeToChange.Add("IsSecured");
            }
            if (originalAttributeMetadata.IsAuditEnabled.Value != readAttributeMetadata.IsAuditEnabled.Value)
            {
                originalAttributeMetadata.IsAuditEnabled = readAttributeMetadata.IsAuditEnabled;
                attributeToChange.Add("IsAuditEnabled");
            }
            return attributeToChange;

        }
        #endregion



    }
}
