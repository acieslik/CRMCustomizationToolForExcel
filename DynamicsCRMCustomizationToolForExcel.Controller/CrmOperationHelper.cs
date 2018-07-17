using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;
using System.Net;

using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Xml.Linq;
using DynamicsCRMCustomizationToolForExcel.Model;

namespace  DynamicsCRMCustomizationToolForExcel.Controller

{
    public class CrmOperationHelper
    {
        private IOrganizationService _service;

        public IOrganizationService Service
        {
            get { return _service; }
            set { _service = value; }
        }

        public void connect(Uri organizationUri)
        {
            Uri homeRealmUri = null;
            ClientCredentials credentials = new ClientCredentials();
            // set default credentials for OrganizationService
            credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;
            // or
            credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
            OrganizationServiceProxy orgProxy = new OrganizationServiceProxy(organizationUri, homeRealmUri, credentials, null);
            _service = (IOrganizationService)orgProxy;
        }

        public void executeOpertionsCrm(CrmOperation operation)
        {
            try
            {
                if (operation.operationType != CrmOperation.CrmOperationType.error)
                {
                    OrganizationResponse orgResp = Service.Execute(operation.orgRequest);
                    operation.operationSucceded = true;
                    operation.orgResponse = orgResp;
                }
            }
            catch (Exception ex)
            {
                operation.exceptionString = ex.Message;
                operation.operationSucceded = false;
            }


        }

        public void publishRequest()
        {
            if (Service != null)
            {
                PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
                Service.Execute(publishRequest);
            }
        }
        public EntityMetadata RetriveEntityAtrribute(Guid enityId)
        {
            if (Service != null)
            {
                try
                {
                    RetrieveEntityRequest lclAEntityMetaDataRequest = new RetrieveEntityRequest();
                    RetrieveEntityResponse lclEntityMetaDataResponse = null;
                    lclAEntityMetaDataRequest.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships;
                    lclAEntityMetaDataRequest.RetrieveAsIfPublished = true;
                    lclAEntityMetaDataRequest.MetadataId = enityId;
                    lclEntityMetaDataResponse = (RetrieveEntityResponse)Service.Execute(lclAEntityMetaDataRequest);
                    return lclEntityMetaDataResponse.EntityMetadata;
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }


        public EntityMetadata RetriveEntityAtrribute(string entityLocicalName)
        {
            if (Service != null)
            {
                try
                {
                    RetrieveEntityRequest lclAEntityMetaDataRequest = new RetrieveEntityRequest();
                    RetrieveEntityResponse lclEntityMetaDataResponse = null;
                    lclAEntityMetaDataRequest.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships;
                    lclAEntityMetaDataRequest.RetrieveAsIfPublished = true;
                    lclAEntityMetaDataRequest.LogicalName = entityLocicalName;
                    lclEntityMetaDataResponse = (RetrieveEntityResponse)Service.Execute(lclAEntityMetaDataRequest);
                    return lclEntityMetaDataResponse.EntityMetadata;
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }

        public IEnumerable<Solution> GetAllSolution()
        {
            QueryExpression query = new QueryExpression()
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet("solutionid", "uniquename", "publisherid"),
                Criteria = new FilterExpression()
            };
            query.Criteria.AddCondition(new ConditionExpression("ismanaged", ConditionOperator.Equal, false));
            query.Criteria.AddCondition(new ConditionExpression("isvisible", ConditionOperator.Equal, true));
            EntityCollection result = Service.RetrieveMultiple(query);
            return result.Entities.Select(x => new Solution() { SolutionId = (Guid)x.Attributes["solutionid"], SolutionName = (string)x.Attributes["uniquename"], PubblisherId = ((EntityReference)x.Attributes["publisherid"]).Id });
        }


        public IEnumerable<Guid> getAllSolutionEntities(Guid solutionId)
        {
            QueryByAttribute componentQuery = new QueryByAttribute
            {
                EntityName = "solutioncomponent",
                ColumnSet = new ColumnSet("componenttype", "objectid", "solutioncomponentid", "solutionid"),
                Attributes = { "solutionid", "componenttype" },
                Values = { solutionId, 1 },
            };
            EntityCollection result = Service.RetrieveMultiple(componentQuery);
            return result.Entities.Select(x => (Guid)x.Attributes["objectid"]);
        }

        public EntityMetadata[] RetriveAllEntities()
        {
            if (Service != null)
            {
                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
              {
                  EntityFilters = EntityFilters.Entity,
                 //| Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                  RetrieveAsIfPublished = true

              };

                RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)Service.Execute(request);
                return response.EntityMetadata;

            }
            return null;
        }

        public OptionSetMetadataBase[] retrieveAllGlobalOptionSets()
        {
            RetrieveAllOptionSetsRequest retrieveAllOptionSetsRequest =
            new RetrieveAllOptionSetsRequest();

            RetrieveAllOptionSetsResponse retrieveAllOptionSetsResponse =
                (RetrieveAllOptionSetsResponse)Service.Execute(retrieveAllOptionSetsRequest);
            return retrieveAllOptionSetsResponse.OptionSetMetadata;
        }

        public int[] InstalledLanguages()
        {
            if (Service != null)
            {
                RetrieveAvailableLanguagesRequest req = new RetrieveAvailableLanguagesRequest();
                RetrieveAvailableLanguagesResponse resp = (RetrieveAvailableLanguagesResponse)Service.Execute(req);
                return resp.LocaleIds;
            }
            return null;
        }

        public EntityCollection GetForms(int typcode)
        {
            if (Service != null)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "systemform",
                    ColumnSet =  new ColumnSet(new string[] { "name"}),
                    Criteria = new FilterExpression()
                };
                query.Criteria.AddCondition(new ConditionExpression("type", ConditionOperator.Equal, 2));
                query.Criteria.AddCondition(new ConditionExpression("objecttypecode", ConditionOperator.Equal, typcode));
                EntityCollection result = Service.RetrieveMultiple(query);
                return result;
            }
            return null;

        }

        public EntityCollection GetViews(int typcode)
        {
            if (Service != null)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "savedquery",
                    ColumnSet = new ColumnSet(new string[] {  "name" }),
                    Criteria = new FilterExpression()
                };
                //query.Criteria.AddCondition(new ConditionExpression("type", ConditionOperator.Equal, 2));
                query.Criteria.AddCondition(new ConditionExpression("returnedtypecode", ConditionOperator.Equal, typcode));
                EntityCollection result = Service.RetrieveMultiple(query);
                return result;
            }
            return null;

        }
        public Entity GetView(Guid viewid)
        {
            if (Service != null)
            {
                //QueryExpression query = new QueryExpression()
                //{
                //    EntityName = "savedquery",
                //    ColumnSet
                //    Criteria = new FilterExpression()
                //};

                //query.Criteria.AddCondition(new ConditionExpression("savedqueryid", ConditionOperator.Equal, viewid));
                RetrieveUnpublishedRequest req = new RetrieveUnpublishedRequest();
                req.ColumnSet = new ColumnSet(new string[] { "layoutxml", "name", "returnedtypecode", "fetchxml" });
                req.Target = new EntityReference("savedquery",viewid);
                RetrieveUnpublishedResponse result = (RetrieveUnpublishedResponse)Service.Execute(req);
                return result.Entity ;
  
            }
            return null;
        }


        
        public Entity GetForm(Guid formid)
        {
            if (Service != null)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "systemform",
                    ColumnSet = new ColumnSet(new string[] { "formxml", "name" }),
                    Criteria = new FilterExpression()
                };

                query.Criteria.AddCondition(new ConditionExpression("formid", ConditionOperator.Equal, formid));
                EntityCollection result = Service.RetrieveMultiple(query);
                if (result.Entities.Count() > 0)
                {
                    return result.Entities.First();
                }
            }
            return null;
        }

        public IEnumerable<string> GetAttributeOfTheMainForm(int typcode)
        {
            if (Service != null)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "systemform",
                    ColumnSet = new ColumnSet(new string[] { "formxml" }),
                    Criteria = new FilterExpression()
                };
                query.Criteria.AddCondition(new ConditionExpression("type", ConditionOperator.Equal, 2));
                query.Criteria.AddCondition(new ConditionExpression("objecttypecode", ConditionOperator.Equal, typcode));
                EntityCollection result = Service.RetrieveMultiple(query);
                if (result.Entities.Count() > 0)
                {
                    string formxml = result.Entities.First().Attributes["formxml"].ToString();
                    XElement form = XElement.Parse(formxml);
                    return from elem in form.Descendants("control")
                           select elem.Attribute("id").Value;
                }
            }
            return null;

        }


        public Guid GetUserId()
        {
            if (Service != null)
            {
                try
                {
                    return ((WhoAmIResponse)Service.Execute(new WhoAmIRequest())).UserId;
                }
                catch 
                {
                    return Guid.Empty;
                }
            }
            return Guid.Empty;
        }



        public List<Publisher> GetPublishersList()
        {
            if (Service != null)
            {
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "publisher",
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression()
                };

                EntityCollection result = Service.RetrieveMultiple(query);
                List<Publisher> PublisherList = new List<Publisher>();

                foreach (var publisher in result.Entities)
                {
                    if (publisher["customizationprefix"].ToString() != "none")
                    {
                        PublisherList.Add(
                            new Publisher()
                            {
                                FriendlyName = publisher["friendlyname"].ToString(),
                                OrganizationId = ((EntityReference)publisher["organizationid"]).Id,
                                PublisherId = (Guid)publisher["publisherid"],
                                UniqueName = publisher["uniquename"].ToString(),
                                CustomizationPrefix = publisher["customizationprefix"].ToString(),
                                IsReadOnly = (bool)publisher["isreadonly"]
                            });
                    }
                }

                return PublisherList;
            }
            return null;

        }

        public RetrieveRelationshipResponse GetOneToManyRelationship(Guid LookUpId)
        {
            if (Service != null)
            {
                RetrieveRelationshipRequest retrieveOneToManyRequest = new RetrieveRelationshipRequest {
                    MetadataId = LookUpId,
                    RetrieveAsIfPublished = true
                };
                return (RetrieveRelationshipResponse)Service.Execute(retrieveOneToManyRequest);
            }
            return null;
        }
    }
}
