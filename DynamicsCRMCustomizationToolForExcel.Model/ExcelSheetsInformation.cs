using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk.Metadata;
using DynamicsCRMCustomizationToolForExcel.Model.FormXml;
using DynamicsCRMCustomizationToolForExcel.Model.FetchXml;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class ExcelSheetsInformation
    {

        private List<ExcelSheetInfo> excelSheets;
        private string curretSheet;
        public ExcelSheetsInformation()
        {
            excelSheets = new List<ExcelSheetInfo>();
            curretSheet = null;
        }

        public void addSheetAndSetAsCurrent(ExcelSheetInfo sheet, string currenSheetName)
        {
            IEnumerable<ExcelSheetInfo> s = excelSheets.Where(x => x.objectName == currenSheetName);
            if (s.Count() > 0)
            {
                excelSheets.Remove(s.First());
            }
            excelSheets.Add(sheet);
            curretSheet = currenSheetName;

        }

        public ExcelSheetInfo getCurrentSheet()
        {
            if (curretSheet != null)
            {
                IEnumerable<ExcelSheetInfo> current = excelSheets.Where(x => x.objectName == curretSheet);
                if (current.Count() > 0)
                {
                    return current.First();
                }
            }
            return null;
        }

        public void removeSheetByName(string name)
        {
            excelSheets.Remove(excelSheets.Where(x => x.objectName == name).FirstOrDefault());

        }

        public ExcelSheetInfo getSheetByName(string name)
        {
            foreach (ExcelSheetInfo sheet in excelSheets)
            {
                if (sheet.objectName.Equals(name))
                {
                    return sheet;
                }
            }
            return null;
        }

        public void setCurrentSheet(string name)
        {
            foreach (ExcelSheetInfo sheet in excelSheets)
            {
                if (sheet.objectName.Equals(name))
                {
                    curretSheet = sheet.objectName;
                }
            }
        }
    }

    public class ExcelSheetInfo
    {
        public enum ExcelSheetType { entity, attribute, optionSet, form ,view};

        public int language { get; set; }
        public bool ignoreChangeEvent { get; set; }
        public string orgPrefix { get; set; }
        public string objectName { get; set; }
        public string workSheetName { get; set; }
        public ExcelSheetType sheetType { get; set; }
        public Worksheet excelsheet { get; set; }

        public ExcelSheetInfo(ExcelSheetType sheetType, Worksheet excelsheet, string objectName)
        {
            this.ignoreChangeEvent = false;
            this.sheetType = sheetType;
            this.excelsheet = excelsheet;
            this.objectName = objectName;
            this.language = GlobalApplicationData.Instance.currentLanguage;
        }

        public ExcelSheetInfo(ExcelSheetType sheetType, string objectName)
        {
            this.ignoreChangeEvent = false;
            this.sheetType = sheetType;
            this.excelsheet = excelsheet;
            this.objectName = objectName;
            this.language = GlobalApplicationData.Instance.currentLanguage;
        }
    }


    public class OptionSetExcelSheetsInfo : ExcelSheetInfo
    {
        public AttributeMetadata parentAttribute { get; set; }
        public OptionSetMetadata optionData { get; set; }

        public OptionSetExcelSheetsInfo(ExcelSheetType sheetType, Worksheet excelsheet, OptionSetMetadata optionData, AttributeMetadata parentAttribute)
            : base(sheetType, excelsheet, optionData.MetadataId.ToString())
        {
            this.optionData = optionData;
            this.parentAttribute = parentAttribute;
        }

        public OptionSetExcelSheetsInfo(ExcelSheetType sheetType, OptionSetMetadata optionData, AttributeMetadata parentAttribute)
            : base(sheetType, optionData.MetadataId.ToString())
        {
            this.optionData = optionData;
            this.parentAttribute = parentAttribute;
        }
    }

    public class EntityExcelSheetsInfo : ExcelSheetInfo
    {

        public EntityMetadata[] entitiesMetadata { get; set; }

        public EntityExcelSheetsInfo(ExcelSheetType sheetType, Worksheet excelsheet, EntityMetadata[] entitiesMetadata, string objectname)
            : base(sheetType, excelsheet, objectname)
        {
            this.entitiesMetadata = entitiesMetadata;
        }

        public EntityExcelSheetsInfo(ExcelSheetType sheetType, EntityMetadata[] entitiesMetadata, string objectname)
            : base(sheetType, objectname)
        {
            this.entitiesMetadata = entitiesMetadata;
        }

    }

    public class AttributeExcelSheetsInfo : ExcelSheetInfo
    {
        public EntityMetadata entityMedata { get; set; }

        public AttributeExcelSheetsInfo(ExcelSheetType sheetType, Worksheet excelsheet, EntityMetadata entityMetadata)
            : base(sheetType, excelsheet, entityMetadata.LogicalName)
        {
            this.entityMedata = entityMetadata;
        }

        public AttributeExcelSheetsInfo(ExcelSheetType sheetType, Worksheet excelsheet, EntityMetadata entityMetadata, string orgPrefix)
            : base(sheetType, excelsheet, entityMetadata.LogicalName)
        {
            this.entityMedata = entityMedata;
            this.orgPrefix = orgPrefix;
        }

        public AttributeExcelSheetsInfo(EntityMetadata entityMedata, string orgPrefix)
            : base(ExcelSheetType.attribute, entityMedata.LogicalName)
        {
            this.entityMedata = entityMedata;
            this.orgPrefix = orgPrefix;
        }

    }

    public class FormExcelSheetsInfo : ExcelSheetInfo
    {
        public string formXml { get; set; }
        public FormType formObj { get; set; }
        public Guid formId { get; set; }

        public FormExcelSheetsInfo(Guid formId, string formXml, FormType formObj)
            : base(ExcelSheetType.form, formId.ToString())
        {
            this.formId = formId;
            this.formXml = formXml;
            this.formObj = formObj;
        }
    }

    public class ViewExcelSheetsInfo : ExcelSheetInfo
    {
        public string layoutXml { get; set; }
        public string fetchxml { get; set; }
        public savedqueryLayoutxmlGrid viewObj { get; set; }
        public FetchType fetchObj { get; set; }
        public List<ViewsRelationsObj>  relationsList { get; set; }
        public Guid viewId { get; set; }
        public  bool isNew {get; set;}
        public string name { get; set; }
        public string description { get; set; }
        public string entity { get; set; }
        public Guid enitityid { get; set; }

        public ViewExcelSheetsInfo(Guid viewId, string layoutXml, string fetchxml, savedqueryLayoutxmlGrid viewObj, FetchType fetchObj, Guid enitityid,string entity)
            : base(ExcelSheetType.view, viewId.ToString())
        {
            this.fetchxml = fetchxml;
            this.viewId = viewId;
            this.layoutXml = layoutXml;
            this.viewObj = viewObj;
            this.fetchObj = fetchObj;
            this.isNew = false;
            this.enitityid = enitityid;
            this.entity = entity;
        }

        public ViewExcelSheetsInfo(Guid viewId , string name , string description , string entity , Guid enitityid)
            : base(ExcelSheetType.view, viewId.ToString())
        {
            this.fetchxml = string.Empty;
            this.viewId = viewId;
            this.layoutXml = string.Empty;
            this.name = name;
            this.description = description;
            this.entity = entity;
            this.viewObj = new savedqueryLayoutxmlGrid();
            this.fetchObj = new FetchType();
            this.isNew = true;
            this.enitityid = enitityid;
        }
    }

}
