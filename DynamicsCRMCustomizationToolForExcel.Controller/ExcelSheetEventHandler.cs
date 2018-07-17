using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk.Metadata;
using DynamicsCRMCustomizationToolForExcel.Model;


namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    class ExcelSheetEventHandler
    {

        private DocEvents_ActivateEventHandler ActivateSheetEventDelegate;
        private DocEvents_FollowHyperlinkEventHandler FollowHyperlinkEventHandlerDelegate;
        public DocEvents_ChangeEventHandler ChangeEventDelegate;
        private GlobalApplicationData appData;
        private ExcelDataAccess excellData;
        public delegate void optioSetEventHandler(OptionSetMetadata optMetadata, AttributeMetadata parentAttribute);

        private optioSetEventHandler optioSetEventHandlerDelegate;

        public ExcelSheetEventHandler(optioSetEventHandler optioSetEventHandlerDelegate, GlobalApplicationData appData, ExcelDataAccess excellData)
        {
            this.optioSetEventHandlerDelegate = optioSetEventHandlerDelegate;
            this.appData = appData;
            this.excellData = excellData;
        }

        public void addEventsToExcelSheet(Worksheet excelSheet)
        {
            //add event link click (OptionSet)
            FollowHyperlinkEventHandlerDelegate = new DocEvents_FollowHyperlinkEventHandler(FollowHyperlinkEventHandler);
            excelSheet.FollowHyperlink += FollowHyperlinkEventHandlerDelegate;
            //event activate sheet
            ActivateSheetEventDelegate = new DocEvents_ActivateEventHandler(ActivateSheetEvent);
            ((DocEvents_Event)excelSheet).Activate += ActivateSheetEventDelegate;
            ChangeEventDelegate = new DocEvents_ChangeEventHandler(ChangeCellEventHandler);
            excelSheet.Change += ChangeEventDelegate;
        }

        private void ActivateSheetEvent()
        {
            string name;
            ExcelSheetInfo.ExcelSheetType type;
            ExcelSheetInfo excelInfo = appData.eSheetsInfomation.getCurrentSheet();
            string orgPrefix;
            int language;
            if (excelInfo != null)
            {
                if (excellData.readSettingRow(excellData.getCurrentSheet(), out name, out type, out orgPrefix, out language))
                {
                    appData.eSheetsInfomation.setCurrentSheet(name);
                    appData.eSheetsInfomation.getCurrentSheet().orgPrefix = orgPrefix;
                    appData.eSheetsInfomation.getCurrentSheet().language = language;
                }
            }
        }

        private void FollowHyperlinkEventHandler(Hyperlink target)
        {
            //Handle the OptionSet click 
            if (target != null)
            { //to check if the link is an option set
                Guid guid = new Guid(target.ScreenTip);
                ExcelSheetInfo currentSh = appData.eSheetsInfomation.getCurrentSheet();
                if (currentSh != null && currentSh is AttributeExcelSheetsInfo)
                {
                    EntityMetadata etMetadata = ((AttributeExcelSheetsInfo)currentSh).entityMedata;
                    IEnumerable<AttributeMetadata> etMetadataFilter = etMetadata.Attributes.Where(x => x is EnumAttributeMetadata && ((EnumAttributeMetadata)x).OptionSet !=null &&((EnumAttributeMetadata)x).OptionSet.MetadataId == guid);
                    if (etMetadataFilter.Count() > 0)
                    {
                        AttributeMetadata currenAttribute = etMetadataFilter.First();
                        OptionSetMetadata optMetadata = ((EnumAttributeMetadata)currenAttribute).OptionSet;
                        optioSetEventHandlerDelegate(optMetadata,currenAttribute);
                        //GlobalOperations.CreatenNewOptionSetSheet(optMetadata.MetadataId, currenAttribute);
                        //appData.eSheetsInfomation.addSheetAndSetAsCurrent(new OptionSetExcelSheetsInfo(ExcelSheetInfo.ExcelSheetType.optionSet, currentSheet, optMetadata, currenAttribute, optionKey), optionKey);
                    }
                }
                if (currentSh != null && currentSh is EntityExcelSheetsInfo && guid != Guid.Empty)
                {
                    IEnumerable<EntityMetadata> currentEntityWithoutAttributes = appData.allEntities.Where(x => x.MetadataId == guid);
                    if (currentEntityWithoutAttributes.Count() != 1)
                    {
                        return;
                    }
                    GlobalOperations.Instance.CreatenNewAttributesSheet(guid);
                }
            }
        }


        public void ChangeCellEventHandler(Range Target)
        {
            ExcelSheetInfo currentsheet = appData.eSheetsInfomation.getCurrentSheet();
            if (appData.eSheetsInfomation.getCurrentSheet() == null || appData.eSheetsInfomation.getCurrentSheet().ignoreChangeEvent)
            {
                return;
            }
            if (Target.Column == ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1 && appData.eSheetsInfomation.getCurrentSheet().sheetType == ExcelSheetInfo.ExcelSheetType.attribute)
            {
                Utils.unprotectSheet(currentsheet);
                currentsheet.excelsheet.Cells[ExcelColumsDefinition.HEADERCOL + 1][Target.Row] = Target.Text == string.Empty ? string.Empty : string.Format("New Field : {0}", Target.Text);
                Utils.protectSheet(currentsheet);
            }
            if (Target.Column == ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1 && appData.eSheetsInfomation.getCurrentSheet().sheetType == ExcelSheetInfo.ExcelSheetType.entity)
            {
                Utils.unprotectSheet(currentsheet);
                currentsheet.excelsheet.Cells[ExcelColumsDefinition.HEADERCOL + 1][Target.Row] = Target.Text == string.Empty ? string.Empty : string.Format("New entity : {0}", Target.Text);
                Utils.protectSheet(currentsheet);
            }
            if (Target.Column == ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1 && appData.eSheetsInfomation.getCurrentSheet() != null && Target.Row > ExcelColumsDefinition.FIRSTROW)
            {
                Utils.unprotectSheet(currentsheet);
                excellData.changeCellsColorOnTypeChange(Target,true);
                Utils.protectSheet(currentsheet);
            }
            if ((Target.Column == ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1 || Target.Column == ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1 || Target.Column == ExcelColumsDefinition.LOOKUPTARGET + 1) && appData.eSheetsInfomation.getCurrentSheet().sheetType == ExcelSheetInfo.ExcelSheetType.attribute && currentsheet.excelsheet.Cells[Target.Row, ExcelColumsDefinition.ATTRIBUTETYPEEXCELCOL + 1].Text == "Lookup")
            {
                Utils.unprotectSheet(currentsheet);
                String lookUpAttributeName = string.Format("{0}_{1}_{2}_{3}", currentsheet.orgPrefix, currentsheet.excelsheet.Cells[Target.Row, ExcelColumsDefinition.LOOKUPTARGET + 1].Text, currentsheet.objectName, Utils.removeOrgPrefix(currentsheet.excelsheet.Cells[Target.Row, ExcelColumsDefinition.SCHEMANAMEEXCELCOL + 1].Text, currentsheet.objectName));
                currentsheet.excelsheet.Cells[Target.Row, ExcelColumsDefinition.LOOKUPRELATIONSHIPNAME + 1] = lookUpAttributeName ;
                Utils.protectSheet(currentsheet);
            }
            if ((Target.Column == ExcelColumsDefinition.VIEWATTRIBUTEENTITY + 1 ) && appData.eSheetsInfomation.getCurrentSheet().sheetType == ExcelSheetInfo.ExcelSheetType.view )
            {
                Utils.unprotectSheet(currentsheet);
                excellData.ShowRelatedAttributes(Target);
                Utils.protectSheet(currentsheet);
            }

        }
    }
}
