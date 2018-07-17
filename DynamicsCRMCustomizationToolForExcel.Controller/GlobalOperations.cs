using DynamicsCRMCustomizationToolForExcel.Model;
using DynamicsCRMCustomizationToolForExcel.Model.FetchXml;
using DynamicsCRMCustomizationToolForExcel.Model.FormXml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class GlobalOperations
    {
        private static GlobalOperations _Instance;
        public static GlobalOperations Instance
        {
            get
            {
                if (_Instance == null)
                {
                    _Instance = new GlobalOperations();            
                }
                return _Instance;
            }
        }

        private GlobalOperations()
        {
            CRMOpHelper = new CrmOperationHelper();
        }

        public CrmOperationHelper CRMOpHelper = null;
        public ExcelDataAccess ExcelOperations;
    
        public  void LoadOperations()
        {
            IEnumerable<EntityMetadata> entitieslist = CRMOpHelper.RetriveAllEntities();
            GlobalApplicationData.Instance.optionSetData = CRMOpHelper.retrieveAllGlobalOptionSets();
            GlobalApplicationData.Instance.crmInstalledLanguages = CRMOpHelper.InstalledLanguages();
            GlobalApplicationData.Instance.crmPubblishers = CRMOpHelper.GetPublishersList();
            GlobalApplicationData.Instance.currentLanguage = GlobalApplicationData.Instance.crmInstalledLanguages.FirstOrDefault();
            GlobalApplicationData.Instance.crmSolutions = CRMOpHelper.GetAllSolution();
            GlobalApplicationData.Instance.allEntities = entitieslist.OrderBy(p => Utils.getLocalizedLabel(p.DisplayName.LocalizedLabels, GlobalApplicationData.Instance.currentLanguage)).ToArray();
            Solution solution = GlobalApplicationData.Instance.crmSolutions.Where(x => x.SolutionName.Equals("Default", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            string DeafultSolution = solution != null ? solution.SolutionName : GlobalApplicationData.Instance.crmSolutions.First().SolutionName;
            GlobalApplicationData.Instance.currentSolution = solution ?? GlobalApplicationData.Instance.crmSolutions.First();
            FilterBySolution(DeafultSolution);
        }

        public void FilterBySolution(string solutionName)
        {
            Solution solution = GlobalApplicationData.Instance.crmSolutions.Where(x => x.SolutionName.Equals(solutionName)).FirstOrDefault();
            IEnumerable<Guid> solutionlist = CRMOpHelper.getAllSolutionEntities(solution.SolutionId);
            GlobalApplicationData.Instance.currentEnitiesList = GlobalApplicationData.Instance.allEntities.Join(solutionlist, x => x.MetadataId, y => y, (x, y) => x).ToArray();
            GlobalApplicationData.Instance.currentSolution = solution ?? GlobalApplicationData.Instance.crmSolutions.First();
        }


        public string getCurrentSolutionPubblisher()
        {
            Publisher p = GlobalApplicationData.Instance.crmPubblishers.Where(x => GlobalApplicationData.Instance.currentSolution.PubblisherId == x.PublisherId).SingleOrDefault();
            return p != null ? p.CustomizationPrefix : GlobalApplicationData.Instance.crmPubblishers.First().CustomizationPrefix;
        }

        public void RefreshCurrentFormSheet()
        {
            GlobalApplicationData data = GlobalApplicationData.Instance;
            if (data.eSheetsInfomation.getCurrentSheet().sheetType == ExcelSheetInfo.ExcelSheetType.form)
            {
                Entity form = CRMOpHelper.GetForm(Guid.Parse(data.eSheetsInfomation.getCurrentSheet().objectName));
                if (form != null && form.Contains("name") && form.Contains("formxml") && form.Contains("objecttypecode"))
                {
                    string sheetName = string.Format("{0} - {1}", form["name"].ToString(), form["objecttypecode"].ToString());
                    FormType formObj = FormXmlMapper.MapFormXmlToObj(form.Attributes["formxml"].ToString());
                    data.eSheetsInfomation.getCurrentSheet().language = GlobalApplicationData.Instance.currentLanguage;
                    ExcelOperations.RefreshFormSheet(data.eSheetsInfomation.getCurrentSheet());
                }
            }
        }


        public void CreatenNewAttributesSheet(Guid entityId)
        {
            GlobalApplicationData data = GlobalApplicationData.Instance;
            EntityMetadata currentEntity = CRMOpHelper.RetriveEntityAtrribute(entityId);
            string sheetName = Utils.getLocalizedLabel(currentEntity.DisplayName.LocalizedLabels, data.currentLanguage);
            AttributeExcelSheetsInfo attributesSheet = new AttributeExcelSheetsInfo(currentEntity, GlobalOperations.Instance.getCurrentSolutionPubblisher());
            CreatenNewExcelSheet(sheetName, attributesSheet);
            IEnumerable<string> formAttr = CRMOpHelper.GetAttributeOfTheMainForm(currentEntity.ObjectTypeCode.Value);
            ExcelOperations.refreshAttributeSheeet(data.eSheetsInfomation.getCurrentSheet(), currentEntity, GlobalApplicationData.Instance.allEntities, formAttr);
        }

        public void CreatenNewFormSheet(Guid formId)
        {
            Entity form = CRMOpHelper.GetForm(formId);
            if (form != null && form.Contains("name") && form.Contains("formxml") && form.Contains("objecttypecode"))
            {
                GlobalApplicationData data = GlobalApplicationData.Instance;
                string sheetName = string.Format("{0} - {1}", form["name"].ToString(), form["objecttypecode"].ToString());
                FormType formObj = FormXmlMapper.MapFormXmlToObj(form.Attributes["formxml"].ToString());
                CreatenNewExcelSheet(sheetName, new FormExcelSheetsInfo(formId, form.Attributes["formxml"].ToString(), formObj));
                data.eSheetsInfomation.getCurrentSheet().language = GlobalApplicationData.Instance.currentLanguage;
                ExcelOperations.RefreshFormSheet(data.eSheetsInfomation.getCurrentSheet());
            }
        }

        public void CreatenNewViewSheet(Guid viewId, Guid entity)
        {
            Entity view = CRMOpHelper.GetView(viewId);
            EntityMetadata etMetadata = CRMOpHelper.RetriveEntityAtrribute(entity);
            IEnumerable<AttributeMetadata> attr = etMetadata.Attributes.Where(x => x.AttributeType != null && x.AttributeType == AttributeTypeCode.Lookup);
            if (view != null && view.Contains("name") && view.Contains("layoutxml") && view.Contains("returnedtypecode") && view.Contains("fetchxml"))
            {
                GlobalApplicationData data = GlobalApplicationData.Instance;
                string sheetName = string.Format("{0} - {1}", view["name"].ToString(), view["returnedtypecode"].ToString());
                savedqueryLayoutxmlGrid viewObj = FormXmlMapper.MapViewXmlToObj(view.Attributes["layoutxml"].ToString());
                FetchType fatchObj = FormXmlMapper.MapFetchXmlToObj(view.Attributes["fetchxml"].ToString());
                CreatenNewExcelSheet(sheetName, new ViewExcelSheetsInfo(viewId, view.Attributes["layoutxml"].ToString(), view.Attributes["fetchxml"].ToString(), viewObj, fatchObj, entity, etMetadata.LogicalName));
                ViewExcelSheetsInfo currentSheet = ((ViewExcelSheetsInfo)data.eSheetsInfomation.getCurrentSheet());
                currentSheet.language = GlobalApplicationData.Instance.currentLanguage;
                currentSheet.relationsList = ViewXml.GenerateViewRelatedObj(currentSheet);
                ExcelOperations.RefreshViewSheet(data.eSheetsInfomation.getCurrentSheet(), attr);
            }
        }



        public  void CreatenNewViewSheet(string viewName, string viewDescription, Guid entity)
        {
            EntityMetadata etMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(entity);
            IEnumerable<AttributeMetadata> attr = etMetadata.Attributes.Where(x => x.AttributeType != null && x.AttributeType == AttributeTypeCode.Lookup);
            GlobalApplicationData data = GlobalApplicationData.Instance;
            string sheetName = string.Format("{0} - {1}", viewName, etMetadata.LogicalName.ToString());
            CreatenNewExcelSheet(sheetName, new ViewExcelSheetsInfo(Guid.NewGuid(), viewName, viewDescription, etMetadata.LogicalName, entity));
            ViewXml.GenerateNewFechXml((ViewExcelSheetsInfo)data.eSheetsInfomation.getCurrentSheet(), etMetadata);
            ViewExcelSheetsInfo currentSheet = ((ViewExcelSheetsInfo)data.eSheetsInfomation.getCurrentSheet());
            currentSheet.language = GlobalApplicationData.Instance.currentLanguage;
            currentSheet.relationsList = ViewXml.GenerateViewRelatedObj(currentSheet);
            GlobalOperations.Instance.ExcelOperations.RefreshViewSheet(data.eSheetsInfomation.getCurrentSheet(), attr);
        }


        public  void RefreshCurrentViewSheet()
        {
            ViewExcelSheetsInfo currentSheet = ((ViewExcelSheetsInfo)GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet());
            Entity view = GlobalOperations.Instance.CRMOpHelper.GetView(currentSheet.viewId);
            EntityMetadata etMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(currentSheet.enitityid);
            IEnumerable<AttributeMetadata> attr = etMetadata.Attributes.Where(x => x.AttributeType != null && x.AttributeType == AttributeTypeCode.Lookup);
            if (view != null && view.Contains("name") && view.Contains("layoutxml") && view.Contains("returnedtypecode") && view.Contains("fetchxml"))
            {
                GlobalApplicationData data = GlobalApplicationData.Instance;
                string sheetName = string.Format("{0} - {1}", view["name"].ToString(), view["returnedtypecode"].ToString());
                savedqueryLayoutxmlGrid viewObj = FormXmlMapper.MapViewXmlToObj(view.Attributes["layoutxml"].ToString());
                FetchType fatchObj = FormXmlMapper.MapFetchXmlToObj(view.Attributes["fetchxml"].ToString());
                currentSheet.layoutXml = view.Attributes["layoutxml"].ToString();
                currentSheet.fetchxml = view.Attributes["fetchxml"].ToString();
                currentSheet.viewObj = viewObj;
                currentSheet.fetchObj = fatchObj;
                currentSheet.language = GlobalApplicationData.Instance.currentLanguage;
                currentSheet.relationsList = ViewXml.GenerateViewRelatedObj(currentSheet);
                GlobalOperations.Instance.ExcelOperations.RefreshViewSheet(data.eSheetsInfomation.getCurrentSheet(), attr);
            }
        }

        public  EntityMetadata[] GetEntitiesWithAttributes()
        {
            EntityMetadata[] allEntities = new EntityMetadata[GlobalApplicationData.Instance.currentEnitiesList.Count()];
            int index = 0;
            foreach (var entity in GlobalApplicationData.Instance.currentEnitiesList)
            {
                if (entity.MetadataId != null)
                {
                    EntityMetadata ent = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(entity.MetadataId.Value);
                    if (ent != null)
                    {
                        allEntities[index] = ent;
                        index++;
                    }
                }
            }

            return allEntities;
        }

        public void CreatenNewEntitySheet()
        {
            GlobalApplicationData data = GlobalApplicationData.Instance;
            string sheetName = string.Format("Entities List - {0}", data.currentSolution.SolutionName);
            EntityMetadata[] allEntities = GetEntitiesWithAttributes();
            EntityExcelSheetsInfo excelSheet = new EntityExcelSheetsInfo(ExcelSheetInfo.ExcelSheetType.entity, allEntities, sheetName);
            CreatenNewExcelSheet(sheetName, excelSheet);
            GlobalOperations.Instance.ExcelOperations.refreshEntitySheeet(data.eSheetsInfomation.getCurrentSheet(), allEntities);
        }

        public  void CreatenNewOptionSetSheet(OptionSetMetadata optionSetMetadata, AttributeMetadata currentAttribute)
        {
            GlobalApplicationData appData = GlobalApplicationData.Instance;
            string sheetName = optionSetMetadata.MetadataId.ToString();
            OptionSetExcelSheetsInfo excelSheet = new OptionSetExcelSheetsInfo(ExcelSheetInfo.ExcelSheetType.optionSet, optionSetMetadata, currentAttribute);
            CreatenNewExcelSheet(sheetName, excelSheet);
            GlobalOperations.Instance.ExcelOperations.refreshOptionSetSheet(GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet());
        }

        public  void CreatenNewExcelSheet(string sheetName, ExcelSheetInfo excelInfo)
        {
            GlobalApplicationData data = GlobalApplicationData.Instance;
            bool sheetRemoved = false;
            ExcelSheetInfo currentSheet = data.eSheetsInfomation.getSheetByName(excelInfo.objectName);
            if (currentSheet != null)
            {
                data.eSheetsInfomation.setCurrentSheet(excelInfo.objectName);
                sheetRemoved = GlobalOperations.Instance.ExcelOperations.selectSheet(currentSheet.workSheetName, data.eSheetsInfomation.getCurrentSheet());
            }
            if (sheetRemoved)
            {
                data.eSheetsInfomation.removeSheetByName(excelInfo.objectName);
            }
            if (data.eSheetsInfomation.getSheetByName(excelInfo.objectName) == null || sheetRemoved)
            {
                string excelName = GetValidWorksheetName(sheetName != string.Empty ? sheetName : excelInfo.objectName);
                Worksheet sheet = GlobalOperations.Instance.ExcelOperations.createNewWorksheet(excelName);
                excelInfo.excelsheet = sheet;
                excelInfo.workSheetName = excelName;
                data.eSheetsInfomation.addSheetAndSetAsCurrent(excelInfo, excelInfo.objectName);
            }
        }

        private  string GetValidWorksheetName(string name)
        {
            // Worksheet name cannot be longer than 31 characters.
            StringBuilder escapedString;
            if (name.Length <= 31)
            {
                escapedString = new StringBuilder(name);
            }
            else
            {
                escapedString = new StringBuilder(name, 0, 31, 31);
            }
            for (int i = 0; i < escapedString.Length; i++)
            {

                if (escapedString[i] == ':' || escapedString[i] == '\\' || escapedString[i] == '/' || escapedString[i] == '?' || escapedString[i] == '*' || escapedString[i] == '[' || escapedString[i] == ']')
                {
                    escapedString[i] = '_';
                }

            }
            return escapedString.ToString();
        }

        public void RefreshCurrentSheet()
        {
            if (GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet() != null)
            {
                GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().language = GlobalApplicationData.Instance.currentLanguage;
                switch (GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet().sheetType)
                {
                    case ExcelSheetInfo.ExcelSheetType.attribute:
                        {
                            EntityMetadata currentEntity = CRMOpHelper.RetriveEntityAtrribute(((AttributeExcelSheetsInfo)GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet()).entityMedata.LogicalName);
                            ((AttributeExcelSheetsInfo)GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet()).entityMedata = currentEntity;
                            IEnumerable<string> formAttr = CRMOpHelper.GetAttributeOfTheMainForm(currentEntity.ObjectTypeCode.Value);
                            ExcelOperations.refreshAttributeSheeet(GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet(), currentEntity, GlobalApplicationData.Instance.allEntities, formAttr);
                            break;
                        }
                    case ExcelSheetInfo.ExcelSheetType.optionSet:
                        {
                            OptionSetExcelSheetsInfo currentOption = (OptionSetExcelSheetsInfo)GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet();
                            EntityMetadata currentEntity = CRMOpHelper.RetriveEntityAtrribute(currentOption.parentAttribute.EntityLogicalName);

                            IEnumerable<AttributeMetadata> currentOpt = currentEntity.Attributes.Where(x => x is EnumAttributeMetadata && ((EnumAttributeMetadata)x).OptionSet != null && ((EnumAttributeMetadata)x).OptionSet.MetadataId.ToString() == currentOption.objectName);

                            if (currentOpt.Count() > 0)
                            {
                                currentOption.optionData = ((EnumAttributeMetadata)currentOpt.First()).OptionSet;
                                ExcelOperations.refreshOptionSetSheet((ExcelSheetInfo)currentOption);
                            }
                            break;
                        }
                    case ExcelSheetInfo.ExcelSheetType.entity:
                        {
                            ExcelSheetInfo sheet = GlobalApplicationData.Instance.eSheetsInfomation.getCurrentSheet();
                            FilterBySolution(GlobalApplicationData.Instance.currentSolution.SolutionName);
                            EntityMetadata[] allEntities = GetEntitiesWithAttributes();
                            GlobalApplicationData.Instance.allEntities = allEntities.OrderBy(p => Utils.getLocalizedLabel(p.DisplayName.LocalizedLabels, sheet.language)).ToArray();
                            ExcelOperations.refreshEntitySheeet(sheet, allEntities);
                            break;
                        }
                    case ExcelSheetInfo.ExcelSheetType.form:
                        {
                            RefreshCurrentFormSheet();
                            break;
                        }
                    case ExcelSheetInfo.ExcelSheetType.view:
                        {
                            RefreshCurrentViewSheet();
                            break;
                        }
                }
            }
        }


    }
}
