using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk.Messages;
using System.Threading;
using DynamicsCRMCustomizationToolForExcel.Model;

 namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class ExcelDataAccess
    {
        private Application _excelApp;
        private Application excelApp
        {
            get { return _excelApp; }
            set { _excelApp = value; }
        }

        public const char optionSetNumberSeparator = ';';
        public const char optionSetSeparator = '\n';
        public const string voidOptionSetString = "Edit Option Set";
        private string[] attributeType = { "Integer", "Lookup", "Picklist", "Memo", "String", "Boolean", "DateTime", "Double", "Decimal", "Money" };

        public ExcelDataAccess(Application excelApp)
        {
            this.excelApp = excelApp;
        }

        public void printEntityAttributeTableHeader(Worksheet excelSheet)
        {
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.HEADERCOL + 1] = "Attribute";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1] = "SchemaName";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.LOGICALNAMEEXCELCOL + 1] = "LogicalName";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DISPLAYNAMEEXCELCOL + 1] = "DisplayName";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DESCRIPTIONEXCELCOL + 1] = "Description";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1] = "AttributeType";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.REQUIREDLEVELEXCELCOL + 1] = "RequiredLevel";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ADVANCEDFINF + 1] = "IsValidForAdvancedFind";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.SECURED + 1] = "IsSecured";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.AUDITENABLED + 1] = "IsAuditEnabled";

            //String Attributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.STRINGMAXLENGTHCOL + 1] = "(String MaxLength)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.STRINGFORMATCOL + 1] = "(String Format)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.STRINGIMEMODECOL + 1] = "(String ImeMode)";

            //Intgert Atributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.INTEGERFORMATCOL + 1] = "(Integer Format)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.INTEGERMAXVALUECOL + 1] = "(Integer Max Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.INTEGERMINVALUECOL + 1] = "(Integer Min Value)";

            //Memo Attributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MEMOFORMATCOL + 1] = "(Memo Format)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MEMOIMEMODECOL + 1] = "(Memo ImeMode)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MEMOMAXLENGTHCOL + 1] = "(Memo MaxLength)";

            // Date time Attributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DATETIMEFORMATCOL + 1] = "(DateTime Format)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DATETIMEIMEMODECOL + 1] = "(DateTime ImeMode)";

            //decimal Attributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DECIMALMAXVALUE + 1] = "(Decimal Max Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DECIMALMINVALUE + 1] = "(Decimal Min Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DECIMALPRECISION + 1] = "(Decimal Precision)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DECIMALIMEMODE + 1] = "(Decimal ImeMode)";

            //double Attributes Columns
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DOUBLEMAXVALUE + 1] = "(Double Max Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DOUBLEMINVALUE + 1] = "(Double Min Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DOUBLEPRECISION + 1] = "(Double Precision)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.DOUBLEIMEMODE + 1] = "(Double ImeMode)";

            //money Attribute Column
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MONEYPRECISION + 1] = "(money Precision)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MONEYMAXVALUE + 1] = "(money Max Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MONEYMINVALUE + 1] = "(money Min Value)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MONEYIMEMODE + 1] = "(money ImeMode)";

            //Option Set
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.PICKLISTREF + 1] = "(Option Set)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.PICKLISTGLOBAL + 1] = "(Global Option Set)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.PICKLISTDEFAULTVALUE + 1] = "(OptionSet Deafult)";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.PICKLISTGLOBALNAME + 1] = "(OptionSet Global Name)";

            //Boolean Attribute Column
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.BOOLEANTRUEOPTION + 1] = "(Boolean true option )";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.BOOLEANFALSEOPTION + 1] = "(Boolean false option )";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.BOOLEANDEFAULTVALUE + 1] = "(Boolean Deafult)";

            //Boolean Attribute Column
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.LOOKUPTARGET + 1] = "(Lookup target )";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1] = "(Lookup attribute name)";

            setHeaderRow(excelSheet);
        }


        public void printEntityTableHeader(Worksheet excelSheet)
        {
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYHEADERCOL + 1] = "Entity";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYLOGICALNAMEEXCELCOL + 1] = "Logical Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYSCHEMANAMEEXCELCOL + 1] = "Schema Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYDISPLAYNAMEEXCELCOL + 1] = "Display Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYPLURALNAME + 1] = "Plural Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYDESCRIPTIONEXCELCOL + 1] = "Description";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYOWNERSHIP + 1] = "Ownership Type";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYDEFINEACTIVITY + 1] = "Activity";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYTYPECODE + 1] = "Object Type Code";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYNOTE + 1] = "Note";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYACTIVITIES + 1] = "activities";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYCONNECTION + 1] = "ConnectionsEnabled";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYSENDMAIL + 1] = "Activity Party";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYMAILMERGE + 1] = "Mail Merge Enabled";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYDOCUMENTMANAGEMENT + 1] = "Document Management Enabled";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYQUEUES + 1] = "Is Valid For Queue";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYMOVEOWNERDEFAULTQUEUE + 1] = "Auto Route To Owner Queue";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYDUPLICATEDETECTION + 1] = "Duplicate Detection Enabled";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYAUDITING + 1] = "Audit Enabled";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYMOBILEEXPRESS + 1] = "Visible In Mobile";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYOFFLINEOUTLOOK + 1] = "Available Offline";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDISPLAYNAME + 1] = "Primary Attribute Display Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTENAME + 1] = "Primary Attribute Schema Name";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDESCRIPTION + 1] = "Primary Attribute Description";
            excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL + 1] = "Primary Attribute Required Level";
        }

        public void setGlobalOptionExcelSheet(Worksheet excelSheet, int numberOfAttributes)
        {
            string[] strBool = { "True", "False" };
            Range range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + 1, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1]);
            setExcelValidationDropDown(range, attributeType, ExcelColumsDefinition.VALIDATIONATTRIBUTETYPE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1], (object)excelSheet.Cells[numberOfAttributes + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1]);
            range.Locked = true;
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1], (object)excelSheet.Cells[numberOfAttributes + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1]);
            range.Locked = true;
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ADVANCEDFINF + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + 1, ExcelColumsDefinition.ADVANCEDFINF + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.SECURED + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + 1, ExcelColumsDefinition.SECURED + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.AUDITENABLED + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + 1, ExcelColumsDefinition.AUDITENABLED + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.REQUIREDLEVELEXCELCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + 1, ExcelColumsDefinition.REQUIREDLEVELEXCELCOL + 1]);
            setExcelValidationDropDown(range, (Enum.GetNames(typeof(AttributeRequiredLevel))).Where(x => !x.Equals("SystemRequired")).ToArray(), ExcelColumsDefinition.VALIDATIONREQUIREDLEVEL, excelSheet);

            // look first 2 columns
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.HEADERCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE - 1, ExcelColumsDefinition.LOGICALNAMEEXCELCOL + 1]);
            range.Locked = true;
            // look header row
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.HEADERCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.MAXNUMBEROFCOLUMN]);
            range.Locked = true;
        }


        public void setSpecificTypeOptionExcelSheet(Worksheet excelSheet, int numberOfAttributes)
        {
            //Option Set
            string[] strBool = { "True", "False" };
            Range range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.PICKLISTGLOBAL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.PICKLISTGLOBAL + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.PICKLISTGLOBALNAME + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.PICKLISTGLOBALNAME + 1]);
            setExcelValidationDropDown(range, GlobalApplicationData.Instance.optionSetData.Select(x => x.Name).ToArray(), ExcelColumsDefinition.VALIDATIONOPTION, excelSheet);
            //Boolean
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.BOOLEANDEFAULTVALUE + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.BOOLEANDEFAULTVALUE + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            //Lookup
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.LOOKUPTARGET + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + +ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.LOOKUPTARGET + 1]);
            setExcelValidationDropDown(range, GlobalApplicationData.Instance.allEntities.Select(x => x.LogicalName).ToArray(), ExcelColumsDefinition.VALIDATIONLOOKUPENTITY, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.LOOKUPTARGET + 1], (object)excelSheet.Cells[numberOfAttributes + 2, ExcelColumsDefinition.LOOKUPTARGET + 1]);
            range.Locked = true;
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1], (object)excelSheet.Cells[numberOfAttributes + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1]);
            range.Locked = true;
            // string
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.STRINGFORMATCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.STRINGFORMATCOL + 1]);
            setExcelValidationDropDown(range, Enum.GetNames(typeof(StringFormat)), ExcelColumsDefinition.VALIDATIONSTRINGFORMAT, excelSheet);
        }


        public void colorAttributeInForm(Range r, IEnumerable<string> formAttributes, string logicalname)
        {
            if (formAttributes != null)
            {
                if (formAttributes.Where(x => x == logicalname).Count() > 0)
                {
                    System.Drawing.Color cellcolor = System.Drawing.Color.LightSteelBlue;
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                }
            }
        }

        public void setSpecificTypeOptionEntityExcelSheet(Worksheet excelSheet, int numberOfAttributes)
        {
            //Option Set
            string[] strBool = { "True", "False" };
            // Booolean
            Range range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYDEFINEACTIVITY + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYDEFINEACTIVITY + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYNOTE + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYNOTE + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYACTIVITIES + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYACTIVITIES + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYCONNECTION + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYCONNECTION + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYSENDMAIL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYSENDMAIL + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYMAILMERGE + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYMAILMERGE + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYDOCUMENTMANAGEMENT + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYDOCUMENTMANAGEMENT + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYQUEUES + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYQUEUES + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYMOVEOWNERDEFAULTQUEUE + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYMOVEOWNERDEFAULTQUEUE + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYDUPLICATEDETECTION + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYDUPLICATEDETECTION + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYAUDITING + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYAUDITING + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYOFFLINEOUTLOOK + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYOFFLINEOUTLOOK + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYMOBILEEXPRESS + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYMOBILEEXPRESS + 1]);
            setExcelValidationDropDown(range, strBool, ExcelColumsDefinition.VALIDATIONTRUEFALSE, excelSheet);
            //ENTITYOWNERSHIP
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL + 1]);
            setExcelValidationDropDown(range, (Enum.GetNames(typeof(AttributeRequiredLevel))).Where(x => !x.Equals("SystemRequired")).ToArray(), ExcelColumsDefinition.VALIDATIONREQUIREDLEVEL, excelSheet);
            //ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL
            range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.ENTITYOWNERSHIP + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYOWNERSHIP + 1]);
            setExcelValidationDropDown(range, Enum.GetNames(typeof(OwnershipTypes)), ExcelColumsDefinition.VALIDATIONENTITYOWNERSHIP, excelSheet);
        }

        private void addLinkEntiy(Worksheet excelSheet, Range cell, EntityMetadata item)
        {
            ExcelSheetInfo sheet = GlobalApplicationData.Instance.eSheetsInfomation.getSheetByName(item.LogicalName);
            String sheetLink = sheet != null ? sheet.excelsheet.Name + "!A2" : "'" + excelSheet.Name + "'!A2";
            string entityString = string.Format("{0} - ({1})", Utils.getLocalizedLabel(item.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance)), item.SchemaName);
            Guid entityguid = (item.MetadataId != null && sheet == null) ? item.MetadataId.Value : Guid.Empty;
            excelSheet.Hyperlinks.Add(
                cell,
                string.Empty,
                 sheetLink,
                 entityguid.ToString(),
                entityString);
        }


        public void printEntityInformation(Worksheet excelSheet, EntityMetadata[] ent)
        {
            if (ent != null)
            {
                printEntityTableHeader(excelSheet);
                setSpecificTypeOptionEntityExcelSheet(excelSheet, ent.Length);
                int excelrow = ExcelColumsDefinition.FIRSTROW;
                foreach (var item in ent.OrderBy(x => Utils.getLocalizedLabel(x.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance))))
                {
                    addLinkEntiy(excelSheet, excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYHEADERCOL + 1], item);
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYLOGICALNAMEEXCELCOL + 1] = item.LogicalName;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYSCHEMANAMEEXCELCOL + 1] = item.SchemaName;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYDISPLAYNAMEEXCELCOL + 1] = Utils.getLocalizedLabel(item.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance));
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYPLURALNAME + 1] = Utils.getLocalizedLabel(item.DisplayCollectionName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance));
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYDESCRIPTIONEXCELCOL + 1] = Utils.getLocalizedLabel(item.Description.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance));
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYOWNERSHIP + 1] = item.OwnershipType != null ? item.OwnershipType.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYDEFINEACTIVITY + 1] = item.IsActivity != null ? item.IsActivity.Value.ToString() : string.Empty;//????????????
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYTYPECODE + 1] = item.ObjectTypeCode != null ? item.ObjectTypeCode.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYNOTE + 1] = /*item. != null ? item.IsActivity.ToString() :*/ string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYACTIVITIES + 1] = /*item.IsActivity != null ? item.IsActivity.ToString() :*/ string.Empty;//?????????----------------
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYCONNECTION + 1] = item.IsConnectionsEnabled != null ? item.IsConnectionsEnabled.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYSENDMAIL + 1] = item.IsActivityParty != null ? item.IsActivityParty.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYMAILMERGE + 1] = item.IsMailMergeEnabled != null ? item.IsMailMergeEnabled.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYDOCUMENTMANAGEMENT + 1] = item.IsDocumentManagementEnabled != null ? item.IsDocumentManagementEnabled.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYQUEUES + 1] = item.IsValidForQueue != null ? item.IsValidForQueue.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYMOVEOWNERDEFAULTQUEUE + 1] = item.AutoRouteToOwnerQueue != null ? item.AutoRouteToOwnerQueue.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYDUPLICATEDETECTION + 1] = item.IsDuplicateDetectionEnabled != null ? item.IsDuplicateDetectionEnabled.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYAUDITING + 1] = item.IsAuditEnabled != null ? item.IsAuditEnabled.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYMOBILEEXPRESS + 1] = item.IsVisibleInMobile != null ? item.IsVisibleInMobile.Value.ToString() : string.Empty;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYOFFLINEOUTLOOK + 1] = item.IsAvailableOffline != null ? item.IsAvailableOffline.Value.ToString() : string.Empty;

                    //AttributeMetadata attr = item.Attributes.Where(x => x.LogicalName == item.PrimaryNameAttribute).FirstOrDefault();
                    //if (attr != null)
                    //{
                    //    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDISPLAYNAME + 1] = Utils.getLocalizedLabel(attr.DisplayName.LocalizedLabels, Utils.getSheetLangauge(appData));
                    //    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTENAME + 1] = attr.SchemaName;
                    //    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDESCRIPTION + 1] = Utils.getLocalizedLabel(attr.Description.LocalizedLabels, Utils.getSheetLangauge(appData));
                    //    excelSheet.Cells[excelrow, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL + 1] = attr.RequiredLevel != null ? attr.RequiredLevel.Value.ToString() : string.Empty;
                    //}
                    excelrow++;
                }
                Range range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.HEADERCOL + 1], (object)excelSheet.Cells[ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.LOGICALNAMEEXCELCOL + 1]);
                range.Locked = true;
                range = excelSheet.get_Range((object)excelSheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.HEADERCOL + 1], (object)excelSheet.Cells[ent.Count() + ExcelColumsDefinition.FIRSTROW - 1, ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL + 1]);
                range.Locked = true;
                formatTable(excelSheet, ExcelColumsDefinition.ENTITYMAXCOLUMNS);
            }
        }

        public void printEntityAttribute(Worksheet excelSheet, AttributeMetadata[] attributeEntity, OneToManyRelationshipMetadata[] relationship, IEnumerable<string> formAttributes)
        {
            if (attributeEntity != null)
            {
                printEntityAttributeTableHeader(excelSheet);
                setGlobalOptionExcelSheet(excelSheet, attributeEntity.Length);
                setSpecificTypeOptionExcelSheet(excelSheet, attributeEntity.Length);
                int excelrow = ExcelColumsDefinition.FIRSTROW;
                foreach (var item in attributeEntity.OrderBy(x => Utils.getLocalizedLabel(x.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance))))
                {
                    colorAttributeInForm(excelSheet.Cells[excelrow, ExcelColumsDefinition.HEADERCOL + 1], formAttributes, item.LogicalName);
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.HEADERCOL + 1] = string.Format("{0} - ({1})", Utils.getLocalizedLabel(item.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance)), item.SchemaName);
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.LOGICALNAMEEXCELCOL + 1] = item.LogicalName;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1] = item.SchemaName;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.DISPLAYNAMEEXCELCOL + 1] = Utils.getLocalizedLabel(item.DisplayName.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance));
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.DESCRIPTIONEXCELCOL + 1] = Utils.getLocalizedLabel(item.Description.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance));
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1] = item.AttributeType.ToString();
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.REQUIREDLEVELEXCELCOL + 1] = item.RequiredLevel.Value.ToString();
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.REQUIREDLEVELEXCELCOL + 1].Locked = !item.RequiredLevel.CanBeChanged;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ADVANCEDFINF + 1] = "'" + item.IsValidForAdvancedFind.Value.ToString();
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.ADVANCEDFINF + 1].Locked = !item.IsValidForAdvancedFind.CanBeChanged;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.SECURED + 1] = "'" + item.IsSecured.Value.ToString();
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.SECURED + 1].Locked = !item.IsCustomAttribute;
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.AUDITENABLED + 1] = "'" + item.IsAuditEnabled.Value.ToString();
                    excelSheet.Cells[excelrow, ExcelColumsDefinition.AUDITENABLED + 1].Locked = !item.IsAuditEnabled.CanBeChanged;
                    printSpecificEntityTypeAttribute(excelSheet.Rows[excelrow], item, relationship);
                    changeCellsColorOnTypeChange(excelSheet.Cells[excelrow, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1], false);
                    excelrow++;

                }
                formatTable(excelSheet, ExcelColumsDefinition.MAXNUMBEROFCOLUMN);
            }
        }


        public void printSettingRow(ExcelSheetInfo excelData, EntityMetadata[] entities)
        {
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGTYPEHEADERCOL] = "Type";
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGTYPECOL] = excelData.sheetType.ToString();
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGNAMEHEADERCOL] = "Name";
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGNAMECOL] = excelData.objectName;
            if (excelData.sheetType == ExcelSheetInfo.ExcelSheetType.attribute && entities != null)
            {
                setExcelValidationDropDown(excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGNAMECOL], entities.Select(x => x.LogicalName).ToArray(), ExcelColumsDefinition.VALIDATIONENTITIES, excelData.excelsheet);
            }
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGLANGUAGEHEADERCOL] = "Language";
            excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGPREFIXHEADERCOL] = "Organization";
            if (GlobalApplicationData.Instance.crmPubblishers.Count() > 0)
            {
                excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGPREFIXPCOL] = excelData.orgPrefix;
                setExcelValidationDropDown(excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGPREFIXPCOL], GlobalApplicationData.Instance.crmPubblishers.Select(x => x.CustomizationPrefix).ToArray(), ExcelColumsDefinition.VALIDATIONPUBBLISHER, excelData.excelsheet);
            }
            if (GlobalApplicationData.Instance.crmInstalledLanguages.Count() > 0)
            {
                excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGLANGUAGECOL] = excelData.language;
                setExcelValidationDropDown(excelData.excelsheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGLANGUAGECOL], GlobalApplicationData.Instance.crmInstalledLanguages.Select(x => x.ToString()).ToArray(), ExcelColumsDefinition.VALIDATIONLANGUAGES, excelData.excelsheet);
            }
            excelData.excelsheet.Rows[ExcelColumsDefinition.SETTINGROW].Hidden = true;
        }

        public bool readSettingRow(Worksheet excelSheet, out string name, out ExcelSheetInfo.ExcelSheetType type, out string orgPrefix, out int language)
        {
            name = string.Empty;
            type = ExcelSheetInfo.ExcelSheetType.attribute;
            orgPrefix = string.Empty;
            language = 1033;
            if (!Enum.IsDefined(typeof(ExcelSheetInfo.ExcelSheetType), excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGTYPECOL].Text)) return false;
            type = (ExcelSheetInfo.ExcelSheetType)Enum.Parse(typeof(ExcelSheetInfo.ExcelSheetType), excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGTYPECOL].Text, true);
            if (excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGNAMECOL].Text == string.Empty) return false;
            name = excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGNAMECOL].Text;
            orgPrefix = excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGPREFIXPCOL].Text;
            int intLanguage;
            if (!int.TryParse(excelSheet.Cells[ExcelColumsDefinition.SETTINGROW, ExcelColumsDefinition.SETTINGLANGUAGECOL].Text, out intLanguage))
            {
                return false;
            }
            language = intLanguage;
            return true;
        }


        public void printSpecificEntityTypeAttribute(Range row, AttributeMetadata attributeEntity, OneToManyRelationshipMetadata[] relationship)
        {
            switch (attributeEntity.AttributeType)
            {
                case AttributeTypeCode.Integer:
                    row.Columns[ExcelColumsDefinition.INTEGERFORMATCOL + 1] = ((IntegerAttributeMetadata)attributeEntity).Format != null ? ((IntegerAttributeMetadata)attributeEntity).Format.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.INTEGERFORMATCOL + 1].Locked = true;
                    row.Columns[ExcelColumsDefinition.INTEGERMAXVALUECOL + 1] = ((IntegerAttributeMetadata)attributeEntity).MaxValue != null ? ((IntegerAttributeMetadata)attributeEntity).MaxValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.INTEGERMINVALUECOL + 1] = ((IntegerAttributeMetadata)attributeEntity).MinValue != null ? ((IntegerAttributeMetadata)attributeEntity).MinValue.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Boolean:
                    row.Columns[ExcelColumsDefinition.BOOLEANTRUEOPTION + 1] = ((BooleanAttributeMetadata)attributeEntity).OptionSet.TrueOption != null ? Utils.getLocalizedLabel(((BooleanAttributeMetadata)attributeEntity).OptionSet.TrueOption.Label.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance)) : string.Empty;
                    row.Columns[ExcelColumsDefinition.BOOLEANFALSEOPTION + 1] = ((BooleanAttributeMetadata)attributeEntity).OptionSet.FalseOption != null ? Utils.getLocalizedLabel(((BooleanAttributeMetadata)attributeEntity).OptionSet.FalseOption.Label.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance)) : string.Empty;
                    row.Columns[ExcelColumsDefinition.BOOLEANDEFAULTVALUE + 1] = ((BooleanAttributeMetadata)attributeEntity).DefaultValue != null ? "'" + ((BooleanAttributeMetadata)attributeEntity).DefaultValue.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.DateTime:
                    row.Columns[ExcelColumsDefinition.DATETIMEFORMATCOL + 1] = ((DateTimeAttributeMetadata)attributeEntity).Format != null ? ((DateTimeAttributeMetadata)attributeEntity).Format.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DATETIMEIMEMODECOL + 1] = ((DateTimeAttributeMetadata)attributeEntity).ImeMode != null ? ((DateTimeAttributeMetadata)attributeEntity).ImeMode.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.String:
                    row.Columns[ExcelColumsDefinition.STRINGMAXLENGTHCOL + 1] = ((StringAttributeMetadata)attributeEntity).MaxLength != null ? ((StringAttributeMetadata)attributeEntity).MaxLength.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.STRINGFORMATCOL + 1] = ((StringAttributeMetadata)attributeEntity).Format != null ? ((StringAttributeMetadata)attributeEntity).Format.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.STRINGFORMATCOL + 1].Locked = true;
                    row.Columns[ExcelColumsDefinition.STRINGIMEMODECOL + 1] = ((StringAttributeMetadata)attributeEntity).ImeMode != null ? ((StringAttributeMetadata)attributeEntity).ImeMode.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Picklist:
                    PicklistAttributeMetadata pkListMetadata = (PicklistAttributeMetadata)attributeEntity;
                    if (pkListMetadata.OptionSet != null)
                    {
                        addOptionSetHyperlink(row.Worksheet, row.Columns[ExcelColumsDefinition.PICKLISTREF + 1], pkListMetadata.OptionSet.MetadataId.ToString(), pkListMetadata.OptionSet.Options);
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBAL + 1] = "'" + pkListMetadata.OptionSet.IsGlobal;
                        row.Columns[ExcelColumsDefinition.PICKLISTDEFAULTVALUE + 1] = pkListMetadata.DefaultFormValue != null ? pkListMetadata.DefaultFormValue.Value.ToString() : string.Empty;
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBALNAME + 1] = pkListMetadata.OptionSet.Name;
                    }

                    break;
                case AttributeTypeCode.Memo:
                    row.Columns[ExcelColumsDefinition.MEMOFORMATCOL + 1] = ((MemoAttributeMetadata)attributeEntity).Format != null ? ((MemoAttributeMetadata)attributeEntity).Format.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.MEMOIMEMODECOL + 1] = ((MemoAttributeMetadata)attributeEntity).ImeMode != null ? ((MemoAttributeMetadata)attributeEntity).ImeMode.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.MEMOMAXLENGTHCOL + 1] = ((MemoAttributeMetadata)attributeEntity).MaxLength != null ? ((MemoAttributeMetadata)attributeEntity).MaxLength.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Double:
                    row.Columns[ExcelColumsDefinition.DOUBLEMAXVALUE + 1] = ((DoubleAttributeMetadata)attributeEntity).MaxValue != null ? "'" + ((DoubleAttributeMetadata)attributeEntity).MaxValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DOUBLEMINVALUE + 1] = ((DoubleAttributeMetadata)attributeEntity).MinValue != null ? "'" + ((DoubleAttributeMetadata)attributeEntity).MinValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DOUBLEPRECISION + 1] = ((DoubleAttributeMetadata)attributeEntity).Precision != null ? ((DoubleAttributeMetadata)attributeEntity).Precision.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DOUBLEIMEMODE + 1] = ((DoubleAttributeMetadata)attributeEntity).ImeMode != null ? ((DoubleAttributeMetadata)attributeEntity).ImeMode.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Decimal:
                    row.Columns[ExcelColumsDefinition.DECIMALMAXVALUE + 1] = ((DecimalAttributeMetadata)attributeEntity).MaxValue != null ? "'" + ((DecimalAttributeMetadata)attributeEntity).MaxValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DECIMALMINVALUE + 1] = ((DecimalAttributeMetadata)attributeEntity).MinValue != null ? "'" + ((DecimalAttributeMetadata)attributeEntity).MinValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DECIMALPRECISION + 1] = ((DecimalAttributeMetadata)attributeEntity).Precision != null ? ((DecimalAttributeMetadata)attributeEntity).Precision.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.DECIMALIMEMODE + 1] = ((DecimalAttributeMetadata)attributeEntity).ImeMode != null ? ((DecimalAttributeMetadata)attributeEntity).ImeMode.Value.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Money:
                    row.Columns[ExcelColumsDefinition.MONEYPRECISION + 1] = ((MoneyAttributeMetadata)attributeEntity).Precision != null ? ((MoneyAttributeMetadata)attributeEntity).Precision.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.MONEYMAXVALUE + 1] = ((MoneyAttributeMetadata)attributeEntity).MaxValue != null ? "'" + ((MoneyAttributeMetadata)attributeEntity).MaxValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.MONEYMINVALUE + 1] = ((MoneyAttributeMetadata)attributeEntity).MinValue != null ? "'" + ((MoneyAttributeMetadata)attributeEntity).MinValue.Value.ToString() : string.Empty;
                    row.Columns[ExcelColumsDefinition.MONEYIMEMODE + 1] = ((MoneyAttributeMetadata)attributeEntity).ImeMode != null ? ((MoneyAttributeMetadata)attributeEntity).ImeMode.ToString() : string.Empty;
                    break;
                case AttributeTypeCode.Lookup:
                    row.Columns[ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1] = getRealationShipName(attributeEntity.LogicalName, relationship);
                    row.Columns[ExcelColumsDefinition.LOOKUPTARGET + 1] = ((LookupAttributeMetadata)attributeEntity).Targets != null && ((LookupAttributeMetadata)attributeEntity).Targets.Count() > 0 ? ((LookupAttributeMetadata)attributeEntity).Targets[0] : string.Empty;
                    break;
                case AttributeTypeCode.State:
                    StateAttributeMetadata stateMetadata = (StateAttributeMetadata)attributeEntity;
                    if (stateMetadata.OptionSet != null)
                    {
                        addOptionSetHyperlink(row.Worksheet, row.Columns[ExcelColumsDefinition.PICKLISTREF + 1], stateMetadata.OptionSet.MetadataId.ToString(), stateMetadata.OptionSet.Options);
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBAL + 1] = string.Empty;
                        row.Columns[ExcelColumsDefinition.PICKLISTDEFAULTVALUE + 1] = stateMetadata.DefaultFormValue != null ? stateMetadata.DefaultFormValue.Value.ToString() : string.Empty;
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBALNAME + 1] = stateMetadata.OptionSet.Name;
                    }
                    break;
                case AttributeTypeCode.Status:
                    StatusAttributeMetadata optMetadata = (StatusAttributeMetadata)attributeEntity;
                    if (optMetadata.OptionSet != null)
                    {
                        addOptionSetHyperlink(row.Worksheet, row.Columns[ExcelColumsDefinition.PICKLISTREF + 1], optMetadata.OptionSet.MetadataId.ToString(), optMetadata.OptionSet.Options);
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBAL + 1] = string.Empty;
                        row.Columns[ExcelColumsDefinition.PICKLISTDEFAULTVALUE + 1] = optMetadata.DefaultFormValue != null ? optMetadata.DefaultFormValue.Value.ToString() : string.Empty;
                        row.Columns[ExcelColumsDefinition.PICKLISTGLOBALNAME + 1] = optMetadata.OptionSet.Name;
                    }
                    break;
            }

        }

        public string getRealationShipName(string attributeName, OneToManyRelationshipMetadata[] relationship)
        {
            IEnumerable<OneToManyRelationshipMetadata> attributeRelationship = relationship.Where(x => x.ReferencingAttribute == attributeName);
            if (attributeRelationship.Count() > 0)
            {
                return attributeRelationship.First().SchemaName;
            }
            return string.Empty;
        }

        public Worksheet checkIfTheSheetExists(string name, Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name.Equals(name))
                {
                    return sheet;
                }
            }
            return null;
        }


        public bool selectSheet(string sheetName, ExcelSheetInfo activeSheet)
        {
            if (checkIfTheSheetExists(sheetName, excelApp.ActiveWorkbook) != null)
            {
                activeSheet.excelsheet.Select(Type.Missing);
                return false;
            }
            return true;
        }

        public Worksheet createNewWorksheet(string sheetName)
        {
            if (excelApp != null)
            {
                Worksheet activeSheet = checkIfTheSheetExists(sheetName, excelApp.ActiveWorkbook);
                if (activeSheet == null)
                {
                    activeSheet = (Worksheet)excelApp.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    activeSheet.Name = sheetName;
                    ExcelSheetEventHandler excelEventHandler = new ExcelSheetEventHandler(createNewOptionSetSheet, GlobalApplicationData.Instance, this);
                    excelEventHandler.addEventsToExcelSheet(activeSheet);
                }
                else
                {
                    createNewWorksheet(sheetName);
                }
                return activeSheet;
            }
            return null;
        }


        public void printValidatioDropDown(Worksheet excelSheet, string[] dropDownValue, int column)
        {
            for (int i = 0; i < dropDownValue.Length; i++)
            {
                excelSheet.Cells[i + 1, column + 1] = dropDownValue[i];
            }
        }
        public void setExcelValidationDropDown(Range cell, string[] dropDownValue, int column, Worksheet excelSheet)
        {
            try
            {
                if (excelSheet.Cells[column + 1][1].Text == string.Empty)
                {
                    printValidatioDropDown(excelSheet, dropDownValue, column);
                }
                cell.Validation.Delete();
                Range r = excelSheet.Range[excelSheet.Cells[column + 1][1], excelSheet.Cells[column + 1][dropDownValue.Length]];
                string reference = new StringBuilder("=").Append(r.get_Address(true, true, XlReferenceStyle.xlA1, true, true)).ToString();
                excelSheet.Names.Add("Range" + column, reference);

                cell.Validation.Add(
                   XlDVType.xlValidateList,
                   XlDVAlertStyle.xlValidAlertInformation,
                   XlFormatConditionOperator.xlBetween,
                   "=Range" + column,
                   Type.Missing);
                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;
            }
            catch (Exception e)
            {
            }
        }

        public void setViewExcelValidationDropDown(Range cell, ViewsRelationsObj obj, int startingcolumn, Worksheet excelSheet, int language)
        {
            try
            {
                int column = startingcolumn;
                if (obj.attributesColumn == null && obj.attributesColumn != 0)
                {
                    while (!string.IsNullOrEmpty(excelSheet.Cells[column + 1][1].Text))
                    {
                        column++;
                    }
                    IEnumerable<string> strData = obj.entityMetadata.Attributes.Where(x => x.IsValidForAdvancedFind != null && x.IsValidForAdvancedFind.Value).Select(x => PrintViewAttributeName(x, language)).OrderBy(x => x);
                    if (excelSheet.Cells[column + 1][1].Text == string.Empty)
                    {
                        printValidatioDropDown(excelSheet, strData.ToArray(), column);
                    }
                    obj.attributesRows = strData.Count();
                    obj.attributesColumn = column;
                }
                column = obj.attributesColumn.Value;
                cell.Validation.Delete();
                Range r = excelSheet.Range[excelSheet.Cells[column + 1][1], excelSheet.Cells[column + 1][obj.attributesRows]];
                string reference = new StringBuilder("=").Append(r.get_Address(true, true, XlReferenceStyle.xlA1, true, true)).ToString();
                excelSheet.Names.Add("Range" + column, reference);
                cell.Validation.Add(
                       XlDVType.xlValidateList,
                       XlDVAlertStyle.xlValidAlertInformation,
                       XlFormatConditionOperator.xlBetween,
                       "=Range" + column,
                       Type.Missing);
                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;
            }
            catch (Exception e)
            {
            }

        }



        public void setHeaderRow(Worksheet sheet)
        {
            excelApp.Application.ActiveWindow.FreezePanes = false;
            sheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.HEADERCOL + 2].Select();
            excelApp.Application.ActiveWindow.FreezePanes = true;
        }

        public void FormatAsTable(Range SourceRange, string TableName, string TableStyleName)
        {
            SourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
            SourceRange, System.Type.Missing, XlYesNoGuess.xlYes, System.Type.Missing).Name =
                TableName;
            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        }

        public void formatTable(Worksheet sheet, int maxcolumn)
        {
            Range SourceRange = (sheet.Range[sheet.Cells[ExcelColumsDefinition.HEADERCOL + 1][ExcelColumsDefinition.HEADERTABLEROW], sheet.Cells[maxcolumn][ExcelColumsDefinition.MAXNUMBEROFATTRIBUTE]]);
            FormatAsTable(SourceRange, "AttributeTable", "TableStyleMedium2");
            sheet.Cells[1][1].Select();
        }

        public void formaOptiontTable(Worksheet sheet)
        {
            Range SourceRange = (sheet.Range[sheet.Cells[ExcelColumsDefinition.HEADERCOL + 1][ExcelColumsDefinition.HEADERTABLEROW], sheet.Cells[ExcelColumsDefinition.MAXNUMBEROFCOLUMNOPTION][ExcelColumsDefinition.MAXNUMBEROFOPTION + 1]]);
            FormatAsTable(SourceRange, "AttributeTable", "TableStyleMedium2");
            sheet.Cells[1][1].Select();
        }

        public ExcelMatrix getExcelDataMatrix(Worksheet excelSheet, int columns, int rows, int emptyColumnToCheck)
        {
            ExcelMatrix excelMatrix = new ExcelMatrix(rows, columns);
            int i = ExcelColumsDefinition.FIRSTROW;
            while (excelSheet.Cells[i, emptyColumnToCheck + 1].Text != string.Empty && i < excelMatrix.rows)
            {
                string[] row = new string[columns];
                for (int j = 1; j <= columns; j++)
                {
                    row[j - 1] = excelSheet.Cells[i, j].Text;
                }
                excelMatrix.AddRow(i - ExcelColumsDefinition.FIRSTROW, row);
                i++;
            }
            return excelMatrix;
        }


        public void addOptionSetHyperlink(Worksheet sheet, Range range, string optionSetGuid, OptionMetadataCollection optionSetCollection)
        {
            bool first = true;
            StringBuilder optBuldier = new StringBuilder();
            foreach (var option in optionSetCollection)
            {
                if (!first)
                    optBuldier.Append(optionSetSeparator);
                else
                    first = false;
                optBuldier.Append(option.Value);
                optBuldier.Append(optionSetNumberSeparator);
                optBuldier.Append(Utils.getLocalizedLabel(option.Label.LocalizedLabels, Utils.getSheetLangauge(GlobalApplicationData.Instance)));
            }
            if (optBuldier.Length == 0)
            {
                optBuldier.Append(voidOptionSetString);
            }
            sheet.Hyperlinks.Add(
                range,
                string.Empty,
                 sheet.Name + "!A2",
                string.Concat(optionSetGuid),
                optBuldier.ToString());
        }

        public void setOptioExcelSettings(Worksheet sheet)
        {
            formaOptiontTable(sheet);
            sheet.Columns.AutoFit();
        }

        public void printOptionSet(ExcelSheetInfo excelData)
        {
            OptionSetMetadata optionSetMetadata = ((OptionSetExcelSheetsInfo)excelData).optionData;
            printSettingRow(excelData, null);
            int i = ExcelColumsDefinition.FIRSTROW;
            excelData.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.OPTIONSETLABELEXCELCOL + 1] = "Option Set Name";
            excelData.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL + 1] = "OptionSet Value";
            foreach (OptionMetadata item in optionSetMetadata.Options)
            {
                excelData.excelsheet.Cells[i, ExcelColumsDefinition.OPTIONSETLABELEXCELCOL + 1] = Utils.getLocalizedLabel(item.Label.LocalizedLabels, excelData.language);
                excelData.excelsheet.Cells[i, ExcelColumsDefinition.OPTIONSETVALUEEXCELCOL + 1] = item.Value;
                i++;
            }
            setOptioExcelSettings(excelData.excelsheet);
        }

        public void createNewOptionSetSheet(OptionSetMetadata optionSetMetadata, AttributeMetadata currentAttribute)
        {
            GlobalOperations.Instance.CreatenNewOptionSetSheet(optionSetMetadata, currentAttribute);

            //GlobalApplicationData data = GlobalApplicationData.Instance;
            //bool sheetRemoved = false;
            //if (data.eSheetsInfomation.getSheetByName(optionSetMetadata.MetadataId.ToString()) != null)
            //{
            //    data.eSheetsInfomation.setCurrentSheet(optionSetMetadata.MetadataId.ToString());
            //    sheetRemoved = Globals.CrmAddIn.excelHelper.selectSheet(data.eSheetsInfomation.getCurrentSheet());
            //}
            //if (sheetRemoved)
            //{
            //    data.eSheetsInfomation.removeSheetByName(optionSetMetadata.MetadataId.ToString());
            //}
            //if (data.eSheetsInfomation.getSheetByName(optionSetMetadata.MetadataId.ToString()) == null || sheetRemoved)
            //{
            //    Worksheet sheet = Globals.CrmAddIn.excelHelper.createNewOptionSetWorksheet(optionSetMetadata.Name);
            //    data.eSheetsInfomation.addSheetAndSetAsCurrent(new OptionSetExcelSheetsInfo(ExcelSheetInfo.ExcelSheetType.optionSet, sheet, optionSetMetadata, data.currentLanguage, currentAttribute), optionSetMetadata.MetadataId.ToString());
            //}
            //data.eSheetsInfomation.getCurrentSheet().language = data.currentLanguage;
            //Globals.CrmAddIn.excelHelper.refreshOptionSetSheet(data.eSheetsInfomation.getCurrentSheet());

        }

        //public Worksheet createNewOptionSetWorksheet(string DisplayName)
        //{
        //    if (excelApp != null)
        //    {
        //        string sheetName = CreateValidWorksheetName(DisplayName);
        //        Worksheet activeSheet = checkIfTheSheetExists(sheetName, excelApp.ActiveWorkbook);
        //        if (activeSheet == null)
        //        {
        //            activeSheet = (Worksheet)excelApp.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //            activeSheet.Name = sheetName;
        //        }
        //        else
        //        {
        //            createNewEntityWorksheet(CreateValidWorksheetName(Guid.NewGuid().ToString()));
        //        }
        //        return activeSheet;
        //    }
        //    return null;

        // if (excelApp != null)
        //{
        //    Worksheet activeSheet = checkIfTheSheetExists(activeSheetName, excelApp.ActiveWorkbook);
        //    if (activeSheet == null)
        //    {
        //        activeSheet = (Worksheet)excelApp.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //        activeSheet.Cells.Locked = false;
        //        activeSheet.Name = activeSheetName;
        //    }
        //    else
        //    {
        //        activeSheet.Select(Type.Missing);
        //    }
        //    return activeSheet;
        //}
        //return null;
        //  }

        public void AutoFitRow(Worksheet excelrange)
        {
            Range rows = excelrange.Rows;
            rows.AutoFit();
            if (rows != null)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rows);
            }
        }


        public Worksheet getCurrentSheet()
        {
            return _excelApp.ActiveSheet;
        }

        public void refreshOptionSetSheet(ExcelSheetInfo sheet)
        {
            sheet.excelsheet.Cells.Clear();
            printOptionSet(sheet);
            sheet.excelsheet.Columns.AutoFit();
        }


        public void refreshAttributeSheeet(ExcelSheetInfo sheetData, EntityMetadata etMatadata, EntityMetadata[] entitiesList, IEnumerable<string> formAttributes)
        {
            GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().ignoreChangeEvent = true;
            Utils.unprotectSheet(sheetData);
            sheetData.excelsheet.Cells.Clear();
            sheetData.excelsheet.Cells.Locked = false;
            printSettingRow(sheetData, entitiesList);
            printEntityAttribute(sheetData.excelsheet, etMatadata.Attributes, etMatadata.ManyToOneRelationships, formAttributes);
            sheetData.excelsheet.Columns.AutoFit();
            sheetData.excelsheet.Columns[ExcelColumsDefinition.HEADERCOL + 1].ColumnWidth = 40;
            Utils.protectSheet(sheetData);
            GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().ignoreChangeEvent = false;
        }


        public void refreshEntitySheeet(ExcelSheetInfo sheetData, EntityMetadata[] entitiesList)
        {
            GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().ignoreChangeEvent = true;
            Utils.unprotectSheet(sheetData);
            sheetData.excelsheet.Cells.Clear();
            sheetData.excelsheet.Cells.Locked = false;
            printSettingRow(sheetData, entitiesList);
            printEntityInformation(sheetData.excelsheet, entitiesList);
            sheetData.excelsheet.Columns.AutoFit();
            sheetData.excelsheet.Columns[ExcelColumsDefinition.HEADERCOL + 1].ColumnWidth = 40;
            Utils.protectSheet(sheetData);
            GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().ignoreChangeEvent = false;
        }

        public bool IsEditing()
        {

            if (excelApp.Interactive != false)
            {
                try
                {
                    excelApp.Interactive = false;
                    excelApp.Interactive = true;
                }
                catch (Exception)
                {
                    return true;
                }
            }
            return false;
        }

        public void changeCellsColorOnTypeChange(Range Target, bool clear)
        {
            Worksheet currentsheet = GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().excelsheet;
            Range r;
            System.Drawing.Color cellcolor = System.Drawing.Color.LightBlue;
            if (clear)
            {
                //r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.MAXNUMBEROFCOLUMN ][Target.Row]];
                //r.Interior.Color = currentsheet.Cells[ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL ][Target.Row].Interior.Color;
            }
            switch ((string)Target.Text)
            {
                case "Integer":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.INTEGERFORMATCOL + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.INTEGERMINVALUECOL + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Boolean":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.BOOLEANTRUEOPTION + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.BOOLEANDEFAULTVALUE + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "DateTime":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.DATETIMEFORMATCOL + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.DATETIMEIMEMODECOL + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "String":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.STRINGMAXLENGTHCOL + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.STRINGIMEMODECOL + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Picklist":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.PICKLISTREF + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.PICKLISTGLOBALNAME + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Memo":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.MEMOFORMATCOL + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.MEMOMAXLENGTHCOL + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Double":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.DOUBLEMAXVALUE + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.DOUBLEIMEMODE + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Decimal":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.DECIMALMAXVALUE + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.DECIMALIMEMODE + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Money":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.MONEYPRECISION + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.MONEYIMEMODE + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                case "Lookup":
                    r = currentsheet.Range[currentsheet.Cells[ExcelColumsDefinition.LOOKUPTARGET + 1][Target.Row], currentsheet.Cells[ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1][Target.Row]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellcolor);
                    break;
                default:
                    break;
            }
        }

        #region view

        public void RefreshViewSheet(ExcelSheetInfo viewSheet, IEnumerable<AttributeMetadata> attr)
        {
            if (viewSheet is ViewExcelSheetsInfo && ((ViewExcelSheetsInfo)viewSheet).viewObj != null)
            {
                ViewExcelSheetsInfo excelview = viewSheet as ViewExcelSheetsInfo;
                try
                {
                    excelview.excelsheet.Cells.Clear();
                    printSettingRow(excelview, null);
                    PrintViewAttributeSheetHeader(excelview);
                    PrintViewAttributes(excelview, attr);
                    formatTable(excelview.excelsheet, ExcelColumsDefinition.VIEWMAXCOLUMNS);
                    excelview.excelsheet.Columns.AutoFit();
                    viewSheet.excelsheet.Columns[ExcelColumsDefinition.VIEWATTRIBUTENAME + 1].ColumnWidth = 40;
                    viewSheet.excelsheet.Columns[ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1].ColumnWidth = 40;
                    AutoFitRow(excelview.excelsheet);
                    excelview.excelsheet.Rows[ExcelColumsDefinition.SETTINGROW].Hidden = true;
                }
                catch (Exception)
                {
                    //--- handle
                }
            }
            else
            {
                //--- handle
            }
        }


        private string getViewEntityName(LookupAttributeMetadata attr, int language)
        {
            string entityDisplayName = string.Empty;
            EntityMetadata etMetadata = GlobalApplicationData.Instance.currentEnitiesList.Where(y => y.LogicalName.Equals(attr.Targets[0], StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (etMetadata != null)
            {
                entityDisplayName = Utils.getLocalizedLabel(etMetadata.DisplayName.LocalizedLabels, language);
            }
            return string.Format("{0} ( {1} ) - {2}.{3}", entityDisplayName,
                        Utils.getLocalizedLabel(attr.DisplayName.LocalizedLabels, language),
                         ((LookupAttributeMetadata)attr).Targets[0],
                         attr.LogicalName);
        }

        private void PrintViewAttributeSheetHeader(ViewExcelSheetsInfo formSheet)
        {
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.VIEWATTRIBUTENAME + 1] = "Attribute";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1] = "Entity";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.VIEWATTRIBUTEDATATYPE + 1] = "Data Type";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.VIEWATTRIBUTEWIDTH + 1] = "Width";
            IEnumerable<ViewsRelationsObj> relObj = formSheet.relationsList.Where(x => string.IsNullOrEmpty(x.relationAlias));
            if (relObj.Count() > 0)
            {
                IEnumerable<string> strinObj = relObj.SelectMany(x => x.entityMetadata.Attributes)
                    .Where(x => x.AttributeType == AttributeTypeCode.Lookup && ((LookupAttributeMetadata)x).Targets.Count() > 0)
                    .Select(x => getViewEntityName(((LookupAttributeMetadata)x), formSheet.language));
                Range range = formSheet.excelsheet.get_Range((object)formSheet.excelsheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1], (object)formSheet.excelsheet.Cells[ExcelColumsDefinition.VIEWMAXATTRIBUTES + 1, ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1]);
                setExcelValidationDropDown(range, strinObj.ToArray(), ExcelColumsDefinition.VALIDATIONTRUEFALSE, formSheet.excelsheet);

                ViewsRelationsObj relObjAttr = (formSheet).relationsList.Where(x => string.IsNullOrEmpty(x.relationAlias)).FirstOrDefault();
                range = formSheet.excelsheet.get_Range((object)formSheet.excelsheet.Cells[ExcelColumsDefinition.FIRSTROW, ExcelColumsDefinition.VIEWATTRIBUTENAME + 1], (object)formSheet.excelsheet.Cells[ExcelColumsDefinition.VIEWMAXATTRIBUTES + 1, ExcelColumsDefinition.VIEWATTRIBUTENAME + 1]);
                setViewExcelValidationDropDown(range, relObjAttr, ExcelColumsDefinition.VIEWVALIDATIONSATTRIBUTES, formSheet.excelsheet, formSheet.language);

            }
        }

        private static string PrintViewAttributeName(AttributeMetadata attr, int language)
        {
            return string.Format("{1} - {0}", attr.LogicalName != null ? attr.LogicalName : string.Empty,
                attr.DisplayName != null ? Utils.getLocalizedLabel(attr.DisplayName.LocalizedLabels, language) : string.Empty);
        }

        public void ShowRelatedAttributes(Range target)
        {
            ExcelSheetInfo currentsheet = GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet();
            if (target.Row >= ExcelColumsDefinition.FIRSTROW && currentsheet is ViewExcelSheetsInfo)
            {
                ViewsRelationsObj relObj = null;
                if (string.IsNullOrEmpty(target.Text))
                {
                    relObj = ((ViewExcelSheetsInfo)currentsheet).relationsList.Where(x => string.IsNullOrEmpty(x.relationAlias)).FirstOrDefault();
                }
                else
                {
                    string[] values = ViewXml.getEntityRelation(target.Text);
                    if (values != null && values.Count() == 2)
                    {
                        relObj = ViewXml.getRelationObj((ViewExcelSheetsInfo)currentsheet, values[0], values[1]);
                    }
                }
                Range cell = currentsheet.excelsheet.get_Range((object)currentsheet.excelsheet.Cells[ExcelColumsDefinition.VIEWATTRIBUTENAME + 1][target.Row], (object)currentsheet.excelsheet.Cells[ExcelColumsDefinition.VIEWATTRIBUTENAME + 1][target.Row]);
                if (relObj != null)
                {
                    setViewExcelValidationDropDown(cell, relObj, ExcelColumsDefinition.VIEWVALIDATIONSATTRIBUTES, currentsheet.excelsheet, currentsheet.language);
                }
            }
        }


        private void PrintViewAttributes(ViewExcelSheetsInfo viewSheet, IEnumerable<AttributeMetadata> attr)
        {
            int langauge = viewSheet.language;
            int rowIndex = ExcelColumsDefinition.FIRSTROW;
            if (viewSheet.viewObj != null && viewSheet.viewObj.row != null && viewSheet.viewObj.row.cell != null)
            {
                foreach (var row in viewSheet.viewObj.row.cell)
                {
                    if (row != null)
                    {
                        IEnumerable<ViewsRelationsObj> currentEt = null;
                        IEnumerable<AttributeMetadata> attMetadata = null;
                        if (!row.name.Contains('.'))
                        {
                            currentEt = viewSheet.relationsList.Where(x => x.relationAlias == null);
                            if (currentEt != null && currentEt.Count() > 0 && currentEt.First() != null && currentEt.First().entityMetadata != null)
                            {
                                attMetadata = currentEt.First().entityMetadata.Attributes.Where(x => x.LogicalName == row.name);
                            }
                        }
                        else
                        {
                            currentEt = viewSheet.relationsList.Where(x => x.relationAlias != null && row.name != null && row.name.StartsWith(x.relationAlias));
                            if (currentEt != null && currentEt.Count() > 0 && currentEt.First() != null && currentEt.First().entityMetadata != null)
                            {
                                attMetadata = currentEt.First().entityMetadata.Attributes.Where(x => string.Concat(currentEt.First().relationAlias + "." + x.LogicalName) == row.name);
                            }
                        }
                        if (currentEt != null && currentEt.Count() > 0 && currentEt.First() != null && currentEt.First().entityMetadata != null && attMetadata != null && attMetadata.Count() > 0)
                        {
                            if (currentEt.First().entityMetadata.PrimaryIdAttribute != attMetadata.First().LogicalName)
                            {
                                viewSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.VIEWATTRIBUTENAME + 1] = attMetadata.Count() > 0 ? PrintViewAttributeName(attMetadata.First(), viewSheet.language) : string.Empty;
                                if (!string.IsNullOrEmpty(currentEt.First().relationAlias))
                                {
                                    ViewsRelationsObj relObj = viewSheet.relationsList.Where(x => string.IsNullOrEmpty(x.relationAlias)).FirstOrDefault();
                                    if (relObj != null)
                                    {
                                        LookupAttributeMetadata lookup = relObj.entityMetadata.Attributes.Where(x => x.AttributeType == AttributeTypeCode.Lookup && x.LogicalName.Equals(currentEt.First().relationTo, StringComparison.InvariantCultureIgnoreCase)).Select(x => (LookupAttributeMetadata)x).FirstOrDefault();
                                        if (lookup != null)
                                        {
                                            viewSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1] = getViewEntityName(lookup, viewSheet.language);
                                        }
                                    }
                                }
                                viewSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.VIEWATTRIBUTEDATATYPE + 1] = attMetadata.Count() > 0 ? attMetadata.First().AttributeType.ToString() : string.Empty;
                                viewSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.VIEWATTRIBUTEWIDTH + 1] = row.width != null ? row.width : string.Empty;
                                rowIndex++;
                            }
                        }
                    }
                }
            }

        }



        #endregion

        #region Forms

        public void RefreshFormSheet(ExcelSheetInfo formSheet)
        {
            if (formSheet is FormExcelSheetsInfo && ((FormExcelSheetsInfo)formSheet).formObj != null)
            {
                FormExcelSheetsInfo excelform = formSheet as FormExcelSheetsInfo;
                try
                {
                    excelform.excelsheet.Cells.Clear();
                    printSettingRow(formSheet, null);
                    PrintFormAttributeSheetHeader(excelform);
                    PrintFormAttributes(excelform);
                    formatTable(excelform.excelsheet, ExcelColumsDefinition.FORMMAXCOLUMNS);
                    excelform.excelsheet.Columns.AutoFit();
                    excelform.excelsheet.Columns[ExcelColumsDefinition.FORMATTRIBUTEEVENTS + 1].ColumnWidth = 80;
                    AutoFitRow(excelform.excelsheet);
                    excelform.excelsheet.Rows[ExcelColumsDefinition.SETTINGROW].Hidden = true;
                }
                catch (Exception)
                {
                    //--- handle
                }
            }
            else
            {
                //--- handle
            }
        }
        private void PrintFormAttributeSheetHeader(FormExcelSheetsInfo formSheet)
        {
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.FORMATTRIBUTELABEL + 1] = "Attribute Label";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.FORMATTRIBUTETAB + 1] = "Tab";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.FORMATTRIBUTESECTION + 1] = "Section";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.FORMATTRIBUTEEVENTS + 1] = "OnChange Events";
            formSheet.excelsheet.Cells[ExcelColumsDefinition.HEADERTABLEROW, ExcelColumsDefinition.FORMATTRIBUTENAME + 1] = "Attribute";
        }

        private void PrintFormAttributes(FormExcelSheetsInfo formSheet)
        {
            int langauge = formSheet.language;
            int rowIndex = ExcelColumsDefinition.FIRSTROW;
            if (formSheet.formObj.tabs != null && formSheet.formObj.tabs.tab != null)
                foreach (var tab in formSheet.formObj.tabs.tab)
                {
                    if (tab != null && tab.columns != null)
                        foreach (var column in tab.columns)
                        {
                            if (column != null && column.sections != null)
                                foreach (var section in column.sections.section)
                                {
                                    if (section != null && section.rows != null)
                                        foreach (var row in section.rows.row)
                                        {
                                            if (row != null && row.cell != null)
                                                foreach (var cell in row.cell)
                                                {
                                                    if (cell != null && cell.control != null)
                                                    {
                                                        try
                                                        {
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTENAME + 1] = cell.control != null && cell.control.id != null ? cell.control.id : string.Empty;
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTELABEL + 1] = cell.labels != null ? FormXmlMapper.GetFormXmlLocalizedLabel(langauge, cell.labels) : string.Empty;
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTETAB + 1] = tab.labels != null ? FormXmlMapper.GetFormXmlLocalizedLabel(langauge, tab.labels) : string.Empty;
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTESECTION + 1] = section.labels != null ? FormXmlMapper.GetFormXmlLocalizedLabel(langauge, section.labels) : string.Empty;
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTEEVENTS + 1].WrapText = true;
                                                            formSheet.excelsheet.Cells[rowIndex, ExcelColumsDefinition.FORMATTRIBUTEEVENTS + 1] = FormXmlMapper.GetAttributeEvents(formSheet.formObj.events, cell.control.id);

                                                            rowIndex++;
                                                        }
                                                        catch (Exception)
                                                        {
                                                            //--- handle
                                                        }
                                                    }
                                                }
                                        }

                                }
                        }
                }
        }

        #endregion

    }
}

