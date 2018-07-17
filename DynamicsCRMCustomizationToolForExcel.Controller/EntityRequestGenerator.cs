using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using DynamicsCRMCustomizationToolForExcel.Model;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class EntityRequestGenerator
    {

        private int languageCode;

        private EntityExcelSheetsInfo entitySheet;

        public EntityRequestGenerator(EntityExcelSheetsInfo entitySheet)
        {
            this.languageCode = entitySheet.language;
            this.entitySheet = entitySheet;
        }


        public IEnumerable<CrmOperation> generateCrmOperationRequest(ExcelMatrix dataMatrix)
        {
            List<CrmOperation> crmOp = new List<CrmOperation>();
            for (int i = 0; i < dataMatrix.numberofElements; i++)
            {
                IEnumerable<EntityMetadata> attribute = entitySheet.entitiesMetadata.Where(x => x.SchemaName == dataMatrix.getElement(i, ExcelColumsDefinition.ENTITYSCHEMANAMEEXCELCOL));
                if (attribute.Count() == 0)
                {
                    addCreateRequest(dataMatrix.getRow(i), crmOp);
                }
                else if (attribute.Count() == 1)
                {
                    //update Entity
                }
            }
            return crmOp;
        }
        private void addSolutionRequest(List<CrmOperation> crmop)
        {
            //AddSolutionComponentRequest addReq = new AddSolutionComponentRequest()
            //{
            //    ComponentType = 
            //    ComponentId = 
            //    SolutionUniqueName = solution.UniqueName
            //};
            //crmop.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.entity, createrequest, strPreview));
        }

        private void addCreateRequest(string[] row, List<CrmOperation> crmop)
        {
            bool parseresult;
            String reqLevelstring = row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEREQUIREMENTLEVEL];
            AttributeRequiredLevel attributeRequiredLevel = AttributeRequiredLevel.None;
            if (Enum.IsDefined(typeof(AttributeRequiredLevel), reqLevelstring))
                attributeRequiredLevel = (AttributeRequiredLevel)Enum.Parse(typeof(AttributeRequiredLevel), reqLevelstring, true);
            CreateEntityRequest createrequest = new CreateEntityRequest
            {
                Entity = new EntityMetadata
                {
                    SchemaName = Utils.addOrgPrefix(row[ExcelColumsDefinition.ENTITYSCHEMANAMEEXCELCOL], entitySheet.orgPrefix, true),
                    DisplayName = new Label(row[ExcelColumsDefinition.ENTITYDISPLAYNAMEEXCELCOL], languageCode),
                    DisplayCollectionName = new Label(row[ExcelColumsDefinition.ENTITYPLURALNAME], languageCode),
                    Description = new Label(row[ExcelColumsDefinition.ENTITYDESCRIPTIONEXCELCOL], languageCode),
                    OwnershipType = Enum.IsDefined(typeof(OwnershipTypes), row[ExcelColumsDefinition.ENTITYOWNERSHIP]) ? (OwnershipTypes)Enum.Parse(typeof(OwnershipTypes), row[ExcelColumsDefinition.ENTITYOWNERSHIP], true) : OwnershipTypes.UserOwned,
                    IsActivity = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYDEFINEACTIVITY], out parseresult) ? parseresult : false,
                    IsConnectionsEnabled = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYCONNECTION], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(true),
                    IsActivityParty = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYSENDMAIL], out parseresult) ? parseresult : false,
                    IsMailMergeEnabled = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYMAILMERGE], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(true),
                    IsDocumentManagementEnabled = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYDOCUMENTMANAGEMENT], out parseresult) ? parseresult : false,
                    IsValidForQueue = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYQUEUES], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(false),
                    AutoRouteToOwnerQueue = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYMOVEOWNERDEFAULTQUEUE], out parseresult) ? parseresult : false,
                    IsDuplicateDetectionEnabled = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYDUPLICATEDETECTION], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(true),
                    IsAuditEnabled = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYAUDITING], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(false),
                    IsVisibleInMobile = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYMOBILEEXPRESS], out parseresult) ? new BooleanManagedProperty(parseresult) : new BooleanManagedProperty(false),
                    IsAvailableOffline = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYOFFLINEOUTLOOK], out parseresult) ? parseresult : false,
                },
                HasNotes = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYNOTE], out parseresult) ? parseresult : true,
                HasActivities = Boolean.TryParse(row[ExcelColumsDefinition.ENTITYACTIVITIES], out parseresult) ? parseresult : true,

                PrimaryAttribute = new StringAttributeMetadata
                {
                    SchemaName = row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDISPLAYNAME] != string.Empty ? row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTENAME] : Utils.addOrgPrefix("name", entitySheet.orgPrefix, true),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(attributeRequiredLevel),
                    MaxLength = 100,
                    Format = StringFormat.Text,
                    DisplayName = row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDISPLAYNAME] != string.Empty ? new Label(row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDISPLAYNAME], languageCode) : new Label("Name", languageCode),
                    Description = new Label(row[ExcelColumsDefinition.ENTITYPRIMARYATTRIBUTEDESCRIPTION], languageCode)
                }

            };
            string strPreview = string.Format("Create new Enity : {0}", createrequest.Entity.SchemaName);
            crmop.Add(new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.entity, createrequest, strPreview));
        }
    }
}
