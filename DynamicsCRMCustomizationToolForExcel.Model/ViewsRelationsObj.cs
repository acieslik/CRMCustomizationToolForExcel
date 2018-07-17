using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class ViewsRelationsObj
    {
        public EntityMetadata entityMetadata{set ; get;}
        public string relationFrom { set; get; }
        public string relationTo { set; get; }
        public string relationAlias { set; get; }
        public string entity { set; get; }
        public int? attributesColumn { set; get; }
        public int? attributesRows { set; get; }
    }
}
