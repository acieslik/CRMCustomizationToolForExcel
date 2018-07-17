using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using DynamicsCRMCustomizationToolForExcel.Model;


namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class Utils
    {
        public const int DEFAULTLANGUAGE=1033;


        public static string addOrgPrefix(string attributeName,string orgPrefix,bool createOperation){
            return !createOperation || attributeName.StartsWith(orgPrefix)  ? attributeName : string.Format("{0}_{1}", orgPrefix, attributeName);
        }

        public static string removeOrgPrefix(string attributeName, string orgPrefix)
        {
            return attributeName.StartsWith(orgPrefix) ? attributeName.Replace(orgPrefix+"_", "") : attributeName;
        }


        public static int getSheetLangauge(GlobalApplicationData appData)
        {
            if (appData.eSheetsInfomation == null || appData.eSheetsInfomation.getCurrentSheet() == null || appData.eSheetsInfomation.getCurrentSheet().language == 0)
            {
                if (appData.crmInstalledLanguages == null || appData.crmInstalledLanguages.Count() == 0 )
                {
                    return DEFAULTLANGUAGE;
                }
                else
                {
                    if (appData.eSheetsInfomation.getCurrentSheet() != null)
                    {
                        appData.eSheetsInfomation.getCurrentSheet().language = appData.crmInstalledLanguages[0];
                    }
                    return appData.crmInstalledLanguages.First();
                }
            }
            return appData.eSheetsInfomation.getCurrentSheet().language;
        }


        public static string getLocalizedLabel(LocalizedLabelCollection labels, int languageCode)
        {
            IEnumerable<LocalizedLabel> label = labels.Where(x => x.LanguageCode == languageCode);
            if (label.Count() == 0)
            {
                return string.Empty ;
            }
            return label.First().Label;
        }

        public static void setLocalizedLabel(LocalizedLabelCollection labels, int languageCode, string labelToSet)
        {
            IEnumerable<LocalizedLabel> label = labels.Where(x => x.LanguageCode == languageCode);
            if (label.Count() != 0)
            {
                label.First().Label = labelToSet;
            }
            else
            {
                labels.Add(new LocalizedLabel(labelToSet, languageCode));
            }
        }

        public static void protectSheet(ExcelSheetInfo currentsheet)
        {
            if (GlobalApplicationData.Instance.enableSheetProtection)
            {
                currentsheet.excelsheet.Protect();
            }
        }

        public static void unprotectSheet(ExcelSheetInfo currentsheet)
        {
            if (GlobalApplicationData.Instance.enableSheetProtection)
            {
                currentsheet.excelsheet.Unprotect();
            }
        }
    }
}
