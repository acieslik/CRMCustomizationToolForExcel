using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
//using CrmCustomizationsExcelAddIn.Helper; //TOCHECK

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class GlobalApplicationData
    {
        private GlobalApplicationData()
        {
            eSheetsInfomation = new ExcelSheetsInformation();
        }

        public ExcelSheetsInformation eSheetsInfomation { get; set; }
        public EntityMetadata[] allEntities { get; set; }
        public EntityMetadata[] currentEnitiesList { get; set; }
        public OptionSetMetadataBase[] optionSetData { get; set; }
        public bool attributeFilterCustom { get; set; }
        public IEnumerable<Solution> crmSolutions { get; set; }
        public IEnumerable<Publisher> crmPubblishers { get; set; }
        public int[] crmInstalledLanguages { get; set; }
        public int currentLanguage { get; set; }
        public Solution currentSolution { get; set; }
        public bool enableSheetProtection { get; set; }
        public bool connectionInProgress { get; set; }
        public Guid selectedEntityTree { get; set; }
        public Guid selectedEntityViewTree { get; set; }

        //TOCHECK
        //public bool IsServiceConnected
        //{
        //    get {
        //        if (Globals.CrmAddIn.crmOpHelper != null && Globals.CrmAddIn.crmOpHelper.Service != null)
        //        {
        //            if (Globals.CrmAddIn.crmOpHelper.GetUserId() != Guid.Empty)
        //            {
        //                return true;
        //            }
        //        }
        //        return false;
        //    }

        //}

        private static GlobalApplicationData instance;

        public static GlobalApplicationData Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new GlobalApplicationData();
                }
                return instance;
            }

        }

    }
}
