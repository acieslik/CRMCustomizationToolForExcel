using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class ExcelColumsDefinition
    {
        //Setting RowData
        public const int SETTINGNAMEHEADERCOL = 1;
        public const int SETTINGNAMECOL = 2;
        public const int SETTINGTYPEHEADERCOL = 3;
        public const int SETTINGTYPECOL = 4;
        public const int SETTINGLANGUAGEHEADERCOL = 5;
        public const int SETTINGLANGUAGECOL = 6;
        public const int SETTINGPREFIXHEADERCOL = 7;
        public const int SETTINGPREFIXPCOL = 8;

        #region attribute entity columns defintaion

        public const int SETTINGROW = 1;
        public const int HEADERTABLEROW = 2;
        public const int FIRSTROW = 3;


        //attribute
        public const int HEADERCOL = 0;
        public const int SCHEMANAMEEXCELCOL = 2;
        public const int LOGICALNAMEEXCELCOL = 1;
        public const int DISPLAYNAMEEXCELCOL = 3;
        public const int DESCRIPTIONEXCELCOL = 4;
        public const int ATTRIBUTETYPEEXCELCOL = 5;
        public const int REQUIREDLEVELEXCELCOL = 6;
        public const int ADVANCEDFINF = 7;
        public const int SECURED = 8;
        public const int AUDITENABLED = 9;

        //String Attributes Columns
        public const int STRINGMAXLENGTHCOL = 10;
        public const int STRINGFORMATCOL = 11;
        public const int STRINGIMEMODECOL = 12;

        //Intgert Atributes Columns
        public const int INTEGERFORMATCOL = 13;
        public const int INTEGERMAXVALUECOL = 14;
        public const int INTEGERMINVALUECOL = 15;

        //Memo Attributes Columns
        public const int MEMOFORMATCOL = 16;
        public const int MEMOIMEMODECOL = 17;
        public const int MEMOMAXLENGTHCOL = 18;

        // Date time Attributes Columns
        public const int DATETIMEFORMATCOL = 19;
        public const int DATETIMEIMEMODECOL = 20;

        //decimal Attributes Columns
        public const int DECIMALMAXVALUE = 21;
        public const int DECIMALMINVALUE = 22;
        public const int DECIMALPRECISION = 23;
        public const int DECIMALIMEMODE = 24;

        //double Attributes Columns
        public const int DOUBLEMAXVALUE = 25;
        public const int DOUBLEMINVALUE = 26;
        public const int DOUBLEPRECISION = 27;
        public const int DOUBLEIMEMODE = 28;

        //money Attribute Column
        public const int MONEYPRECISION = 29;
        public const int  MONEYMAXVALUE= 30;
        public const int MONEYMINVALUE= 31;
        public const int MONEYIMEMODE = 32;

        // OptionSet Attribute Column
        public const int PICKLISTREF = 33;
        public const int PICKLISTGLOBAL = 34;
        public const int PICKLISTDEFAULTVALUE = 35;
        public const int PICKLISTGLOBALNAME = 36;

        // Boolean Attribute Column
        public const int BOOLEANTRUEOPTION = 37;
        public const int BOOLEANFALSEOPTION = 38;
        public const int BOOLEANDEFAULTVALUE = 39;

        // Boolean Attribute Column

        public const int LOOKUPTARGET = 40 ;
        public const int LOOKUPRELATIONSHIPNAME = 41;

        #endregion

        #region optionset attribute definitio

        public const int OPTIONSETLABELEXCELCOL = 0;
        public const int OPTIONSETVALUEEXCELCOL = 1;

        #endregion

        public const int MAXNUMBEROFOPTION = 1024;
        public const int MAXNUMBEROFCOLUMNOPTION = 2;
        public const int MAXNUMBEROFATTRIBUTE = 1024;
        public const int MAXNUMBEROFCOLUMN = 42;


        #region Valitation drop down definition

        public const int VALIDATIONATTRIBUTETYPE = 50;
        public const int VALIDATIONTRUEFALSE = 51;
        public const int VALIDATIONREQUIREDLEVEL = 52;
        public const int VALIDATIONSTRINGFORMAT = 53;
        public const int VALIDATIONOPTION = 54;
        public const int VALIDATIONLOOKUPENTITY = 55;
        public const int VALIDATIONENTITIES = 56;
        public const int VALIDATIONLANGUAGES = 57;
        public const int VALIDATIONPUBBLISHER = 58;
        public const int VALIDATIONENTITYOWNERSHIP = 59;
        #endregion


        #region entity definition


        //attribute
        public const int ENTITYHEADERCOL = 0;
        public const int ENTITYSCHEMANAMEEXCELCOL = 2;
        public const int ENTITYLOGICALNAMEEXCELCOL = 1;
        public const int ENTITYDISPLAYNAMEEXCELCOL = 3;
        public const int ENTITYPLURALNAME = 4;
        public const int ENTITYDESCRIPTIONEXCELCOL = 5;
        public const int ENTITYOWNERSHIP = 6;
        public const int ENTITYDEFINEACTIVITY = 7;
        public const int ENTITYTYPECODE = 8;
        public const int ENTITYNOTE = 9;
        public const int ENTITYACTIVITIES = 10;
        public const int ENTITYCONNECTION = 11;
        public const int ENTITYSENDMAIL = 12;
        public const int ENTITYMAILMERGE = 13;
        public const int ENTITYDOCUMENTMANAGEMENT = 14;
        public const int ENTITYQUEUES = 15;
        public const int ENTITYMOVEOWNERDEFAULTQUEUE = 16;
        public const int ENTITYDUPLICATEDETECTION = 17;
        public const int ENTITYAUDITING = 18;
        public const int ENTITYMOBILEEXPRESS = 19;
       // public const int ENTITYREADINGPANEOUTLOOK= 19;
        public const int ENTITYOFFLINEOUTLOOK = 20;

        //PRIMARY ATTRIBUTE 
        public const int ENTITYPRIMARYATTRIBUTEDISPLAYNAME = 21;
        public const int ENTITYPRIMARYATTRIBUTENAME = 22;
        public const int ENTITYPRIMARYATTRIBUTEDESCRIPTION = 23;
        public const int ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL = 24;

        public const int ENTITYMAXCOLUMNS = 25;
        #endregion

        #region form definition

        public const int FORMATTRIBUTENAME = 0;
        public const int FORMATTRIBUTELABEL = 1;
        public const int FORMATTRIBUTETAB = 2;
        public const int FORMATTRIBUTESECTION = 3;
        public const int FORMATTRIBUTEEVENTS = 4;


        public const int FORMMAXCOLUMNS = 5;
        #endregion

        #region view definition

        public const int VIEWATTRIBUTENAME = 0;
        public const int VIEWATTRIBUTEENTITY= 1;
        public const int VIEWATTRIBUTEDATATYPE= 2;
        public const int VIEWATTRIBUTEWIDTH = 3;

        public const int VIEWMAXCOLUMNS = 4;
        public const int VIEWMAXATTRIBUTES = 100;

        public const int VIEWVALIDATIONSATTRIBUTES = 60;
        #endregion


    }
}
