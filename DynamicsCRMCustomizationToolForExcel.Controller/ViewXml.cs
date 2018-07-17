
using DynamicsCRMCustomizationToolForExcel.Model;
using DynamicsCRMCustomizationToolForExcel.Model.FetchXml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class ViewXml
    {

        #region Views

        public static List<ViewsRelationsObj> GenerateViewRelatedObj(ViewExcelSheetsInfo sheetView)
        {
            List<ViewsRelationsObj> listObj = new List<ViewsRelationsObj>();
            if (sheetView.viewObj != null && sheetView.viewObj.row != null && sheetView.viewObj.row.cell != null)
            {
                foreach (var row in sheetView.viewObj.row.cell)
                {
                    if (row != null)
                    {
                        updateViewRelatedObj(sheetView.fetchObj, row.name, listObj);
                    }
                }
            }
            return listObj;
        }


        public static void GenerateNewFechXml(ViewExcelSheetsInfo sheetView, EntityMetadata etMetadata)
        {
            FetchEntityType et = new FetchEntityType();
            et.name = etMetadata.LogicalName;
            et.Items = new object[3] { new FetchOrderType() { attribute = etMetadata.PrimaryIdAttribute, descending = false }, new FetchAttributeType() { name = etMetadata.PrimaryIdAttribute }, new FetchAttributeType() { name = etMetadata.PrimaryNameAttribute } };
            sheetView.fetchObj.Items = new object[] { et };
            sheetView.fetchObj.version = "1.0";
            sheetView.fetchObj.outputformat = FetchTypeOutputformat.xmlplatform;
            sheetView.fetchObj.mapping = FetchTypeMapping.logical;
            savedqueryLayoutxmlGridRowCell[] cells = new savedqueryLayoutxmlGridRowCell[1] { 
                new savedqueryLayoutxmlGridRowCell() { name = etMetadata.PrimaryNameAttribute, width = "150" }
            };
            sheetView.viewObj.row = new savedqueryLayoutxmlGridRow() { cell = cells, name = "result", id = etMetadata.PrimaryIdAttribute };
            sheetView.viewObj.select = "1";
            sheetView.viewObj.preview = "1";
            sheetView.viewObj.@object = etMetadata.ObjectTypeCode.ToString();
            sheetView.viewObj.icon = "1";
            sheetView.viewObj.jump = etMetadata.PrimaryNameAttribute;
            sheetView.viewObj.name = "resultset";
        }

        public static ViewsRelationsObj getRelationObj(ViewExcelSheetsInfo sheetView, string entity, string fromAttribute)
        {
            ViewsRelationsObj obj = sheetView.relationsList.Where(x => x.relationFrom != null && x.entity.Equals(entity, StringComparison.InvariantCultureIgnoreCase) &&
                                                x.relationFrom.Equals(fromAttribute, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (obj == null)
            {
                ViewsRelationsObj mainobj = sheetView.relationsList.Where(x => x.relationAlias == null).FirstOrDefault();
                OneToManyRelationshipMetadata rel = mainobj.entityMetadata.ManyToOneRelationships
                    .Where(x => x.ReferencedEntity.Equals(entity, StringComparison.InvariantCultureIgnoreCase) &&
                           x.ReferencingAttribute.Equals(fromAttribute, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (rel != null)
                {
                    obj = new ViewsRelationsObj()
                    {
                        entity = entity.ToLower(),
                        entityMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(entity),
                        relationAlias = string.Concat("a_", Guid.NewGuid().ToString().Replace("-", "")).ToLower(),
                        relationFrom = rel.ReferencedAttribute.ToLower(),
                        relationTo = fromAttribute.ToLower()
                    };
                    sheetView.relationsList.Add(obj);
                }
            }
            return obj;
        }


        private static void updateViewRelatedObj(FetchType fetchType, string attributeName, List<ViewsRelationsObj> relationObj)
        {
            int index = attributeName.IndexOf(".");
            if (index == -1)
            {
                if (relationObj.Where(x => x.relationAlias == null).Count() == 0)
                {
                    foreach (object fetch in fetchType.Items)
                    {
                        if (fetch is FetchEntityType)
                        {
                            foreach (object obj in ((FetchEntityType)fetch).Items)
                            {
                                if (obj is FetchAttributeType && attributeName == ((FetchAttributeType)obj).name)
                                {
                                    FetchEntityType fetchent = (FetchEntityType)fetch;
                                    relationObj.Add(new ViewsRelationsObj()
                                    {
                                        entity = fetchent.name,
                                        entityMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(fetchent.name)
                                    });
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                string alias = attributeName.Substring(0, index);
                if (relationObj.Where(x => x.relationAlias != null && attributeName.StartsWith(x.relationAlias)).Count() == 0)
                {
                    foreach (object fetch in fetchType.Items)
                    {
                        foreach (object obj in ((FetchEntityType)fetch).Items)
                        {
                            if (obj is FetchLinkEntityType)
                            {
                                FetchLinkEntityType relation = getFechXmlFromLinks((FetchLinkEntityType)obj, attributeName);
                                if (relation != null)
                                {
                                    relationObj.Add(new ViewsRelationsObj()
                                    {
                                        entity = relation.name,
                                        entityMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(relation.name),
                                        relationAlias = relation.alias,
                                        relationFrom = relation.from,
                                        relationTo = relation.to
                                    });
                                }
                            }
                        }
                    }
                }
            }
        }


        private static FetchLinkEntityType getFechXmlFromLinks(FetchLinkEntityType fetchLinkType, string attributeName)
        {
            if (attributeName.Contains("."))
            {
                if (attributeName.StartsWith(string.Concat(((FetchLinkEntityType)fetchLinkType).alias)))
                {
                    return (FetchLinkEntityType)fetchLinkType;
                }
                if (fetchLinkType != null && fetchLinkType.Items != null && fetchLinkType.Items.Count() > 0)
                {
                    foreach (object fetchType in fetchLinkType.Items)
                    {
                        if (fetchType is FetchLinkEntityType)
                        {
                            return getFechXmlFromLinks((FetchLinkEntityType)fetchType, attributeName);

                        }
                    }
                }
            }
            return null;
        }
        #endregion

        #region Views

        public static void getFechAndLayouXml(ExcelMatrix dataMatrix, ViewExcelSheetsInfo sheetView, out string fetchXml, out string viewXml)
        {
            List<ViewFeo> viewFeoList = new List<ViewFeo>();
            for (int i = 0; i < dataMatrix.numberofElements; i++)
            {

                string[] row = dataMatrix.getRow(i);
                string[] values = getEntityRelation(row[ExcelColumsDefinition.VIEWATTRIBUTEENTITY]);
                string attributeName = row[ExcelColumsDefinition.VIEWATTRIBUTENAME];
                if (attributeName.LastIndexOf("-") >= 0)
                {
                    attributeName = attributeName.Substring(attributeName.LastIndexOf("-") + 1, attributeName.Length - attributeName.LastIndexOf("-") - 1);
                    attributeName = attributeName.Trim();
                }
                string attributeEntity = string.Empty;
                string attributeRelationAttribute = string.Empty;
                if (values != null && values.Count() == 2)
                {
                    attributeEntity = values[0];
                    attributeRelationAttribute = values[1];
                }
                string attributewidth = row[ExcelColumsDefinition.VIEWATTRIBUTEWIDTH];
                ViewsRelationsObj obj = null;
                if (!string.IsNullOrEmpty(attributeEntity))
                {
                    obj = sheetView.relationsList.Where(x => x.relationAlias != null && x.entity.Equals(attributeEntity, StringComparison.InvariantCultureIgnoreCase) &&
                            x.relationTo.Equals(attributeRelationAttribute, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                }
                else
                {
                    obj = sheetView.relationsList.Where(x => x.relationAlias == null).FirstOrDefault();
                    attributeEntity = obj.entity;
                }

                ViewFeo viewFeo = new ViewFeo()
                {
                    AttributeName = attributeName,
                    AttributeEntity = attributeEntity,
                    AttributeWidth = attributewidth,
                    AttributeObj = obj,
                };
                if (obj == null)
                {
                    setAttributesObj(viewFeo, sheetView.relationsList);
                }
                viewFeoList.Add(viewFeo);
            }
            viewXml = getViewXmlString(viewFeoList, sheetView.viewObj);
            removeUnnecessaryLinksAttributes(sheetView.viewObj, sheetView);
            fetchXml = getFecthXmlString(viewFeoList, sheetView.fetchObj);
        }


        public static IEnumerable<CrmOperation> generateCrmOperationRequest(ExcelMatrix dataMatrix, ViewExcelSheetsInfo sheetView)
        {
            List<CrmOperation> crmOp = new List<CrmOperation>();
            try
            {
                string fetchXml;
                string viewXml;
                getFechAndLayouXml( dataMatrix,  sheetView, out  fetchXml, out  viewXml);
                if (sheetView.isNew)
                {
                    crmOp.Add(generateUpdateRequest(viewXml, fetchXml, sheetView.name, sheetView.description, sheetView.entity));
                }
                else
                {
                    crmOp.Add(generateUpdateRequest(viewXml, fetchXml, sheetView.viewId));
                }
            }
            catch (Exception ex)
            {
                crmOp.Add(new CrmOperation(CrmOperation.CrmOperationType.error, CrmOperation.CrmOperationTarget.none, null, "Generic error"));
            }
            return crmOp;
        }

        private static void removeUnnecessaryLinksAttributes(savedqueryLayoutxmlGrid viewFeo, ViewExcelSheetsInfo excelData)
        {
            FetchEntityType currentEt = excelData.fetchObj.Items.Where(x => x is FetchEntityType).Select(x => (FetchEntityType)x).FirstOrDefault();
            ViewsRelationsObj obj = excelData.relationsList.Where(x => x.relationAlias == null).FirstOrDefault();
            if (currentEt != null && obj != null)
            {

                List<object> array = currentEt.Items.ToList();
                for (int i = array.Count - 1; i > -1; i--)
                {
                    object currentItem = array[i];
                    if (currentItem is FetchLinkEntityType)
                    {
                        FetchLinkEntityType link = currentItem as FetchLinkEntityType;
                        removeUnnecessaryAttributesInner(link, viewFeo);
                        if (link == null || link.visible == false)
                        {
                            if (link == null || link.Items.Count() == 0)
                            {
                                array.RemoveAt(i);
                            }
                        }
                    }
                    else if (currentItem is FetchAttributeType)
                    {
                        FetchAttributeType attr = currentItem as FetchAttributeType;
                        if (!(viewFeo.row.cell.Where(x => x.name == attr.name).Count() > 0) && attr.name != obj.entityMetadata.PrimaryIdAttribute)
                        {
                            array.RemoveAt(i);
                        }
                    }
                }
                currentEt.Items = array.ToArray();
            }
        }


        private static void removeUnnecessaryAttributesInner(FetchLinkEntityType link, savedqueryLayoutxmlGrid viewFeo)
        {
            if (link.Items != null)
            {
                List<object> array = link.Items.ToList();
                for (int i = array.Count - 1; i > -1; i--)
                {
                    object currentItem = array[i];
                    if (currentItem is FetchAttributeType)
                    {
                        FetchAttributeType attr = currentItem as FetchAttributeType;
                        string alias = string.Format("{0}.{1}", link.alias, attr.name);
                        if (!(viewFeo.row.cell.Where(x => x.name == attr.name).Count() > 0))
                        {
                           array.RemoveAt(i);
                        }
                    }
                }
                link.Items = array.ToArray();
            }
        }

        private static void setAttributesObj(ViewFeo viewFeo, List<ViewsRelationsObj> relObj)
        {
            ViewsRelationsObj mainobj = relObj.Where(x => x.relationAlias == null).FirstOrDefault();
            OneToManyRelationshipMetadata rel = mainobj.entityMetadata.ManyToOneRelationships.Where(x => x.ReferencedEntity.Equals(viewFeo.AttributeEntity, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (rel != null)
            {

                viewFeo.AttributeObj = new ViewsRelationsObj()
                {
                    entity = viewFeo.AttributeEntity.ToLower(),
                    entityMetadata = GlobalOperations.Instance.CRMOpHelper.RetriveEntityAtrribute(viewFeo.AttributeEntity),
                    relationAlias = string.Concat("a_", Guid.NewGuid().ToString().Replace("-", "")).ToLower(),
                    relationFrom = rel.ReferencedAttribute.ToLower(),
                    relationTo = rel.ReferencingAttribute.ToLower()
                };
                relObj.Add(viewFeo.AttributeObj);
            }
        }


        private static string getViewXmlString(List<ViewFeo> viewFeoList, savedqueryLayoutxmlGrid viewObj)
        {
            List<savedqueryLayoutxmlGridRowCell> newcellList = new List<savedqueryLayoutxmlGridRowCell>();
            if (viewObj.row != null && viewObj.row.cell != null)
            {
                List<savedqueryLayoutxmlGridRowCell> oldcellList = viewObj.row.cell.ToList();
                foreach (ViewFeo viewFeo in viewFeoList)
                {
                    if (viewFeo.AttributeObj != null)
                    {
                        string attributeNameAlias = string.IsNullOrEmpty(viewFeo.AttributeObj.relationAlias) ? viewFeo.AttributeName : string.Concat(viewFeo.AttributeObj.relationAlias, ".", viewFeo.AttributeName);
                        addCellToView(newcellList, oldcellList, attributeNameAlias, viewFeo.AttributeWidth);
                    }
                    else
                    {
                        throw new Exception("Relation Not Found");
                    }
                }
                viewObj.row.cell = newcellList.ToArray();
            }
            return FormXmlMapper.MapObjToViewXml(viewObj);
        }

        private static string getFecthXmlString(List<ViewFeo> viewFeoList, FetchType fecthObj)
        {
            foreach (ViewFeo viewFeo in viewFeoList)
            {
                if (viewFeo.AttributeObj != null)
                {
                    addNewAttributeToFech(fecthObj, viewFeo.AttributeName, viewFeo.AttributeObj);
                }
                else
                {
                    throw new Exception("Relation Not Found");
                }
            }
            return FormXmlMapper.MapObjToFetchXml(fecthObj);
        }


        private static void addNewAttributeToFech(FetchType fecthObj, string attributeName, ViewsRelationsObj currentEntity)
        {
            if (string.IsNullOrEmpty(currentEntity.relationAlias))
            {
                FetchEntityType currentEt = fecthObj.Items.Where(x => x is FetchEntityType && ((FetchEntityType)x).name.Equals(currentEntity.entity, StringComparison.InvariantCultureIgnoreCase)).Select(x => (FetchEntityType)x).FirstOrDefault();
                if (currentEt != null)
                {
                    List<object> items = currentEt.Items.ToList();
                    FetchAttributeType attribute = items.Where(x => x is FetchAttributeType && ((FetchAttributeType)x).name.Equals(attributeName)).Select(x => (FetchAttributeType)x).FirstOrDefault();
                    if (attribute == null)
                    {
                        List<object> objectList = currentEt.Items.ToList();
                        objectList.Add(new FetchAttributeType()
                        {
                            name = attributeName
                        });
                        currentEt.Items = objectList.ToArray();
                    }
                }
            }
            else
            {
                FetchEntityType currentEt = fecthObj.Items.Where(x => x is FetchEntityType).Select(x => (FetchEntityType)x).FirstOrDefault();
                if (currentEt != null)
                {
                    FetchLinkEntityType currentLink = currentEt.Items.Where(x => x is FetchLinkEntityType &&
                        ((FetchLinkEntityType)x).alias.Equals(currentEntity.relationAlias, StringComparison.InvariantCultureIgnoreCase))
                        .Select(x => (FetchLinkEntityType)x).FirstOrDefault();
                    if (currentLink != null)
                    {
                        List<object> items = currentLink.Items.ToList();
                        FetchAttributeType attribute = items.Where(x => x is FetchAttributeType && ((FetchAttributeType)x).name.Equals(attributeName)).Select(x => (FetchAttributeType)x).FirstOrDefault();
                        if (attribute == null)
                        {
                            List<object> objectList = currentLink.Items.ToList();
                            objectList.Add(new FetchAttributeType()
                            {
                                name = attributeName
                            });
                            currentLink.Items = objectList.ToArray();
                        }
                    }
                    else
                    {
                        List<object> objectList = currentEt.Items.ToList();
                        objectList.Add(new FetchLinkEntityType()
                        {
                            Items = new object[] {(new FetchAttributeType()
                            {
                                name = attributeName
                            })},
                            name = currentEntity.entity,
                            alias = currentEntity.relationAlias,
                            linktype = "outer",
                            from = currentEntity.relationFrom,
                            to = currentEntity.relationTo,
                            intersect = false,
                            intersectSpecified = false,
                            visible = false,
                            visibleSpecified = false,
                        });
                        currentEt.Items = objectList.ToArray();
                    }
                }
            }
        }

        private static void addCellToView(List<savedqueryLayoutxmlGridRowCell> cellList, List<savedqueryLayoutxmlGridRowCell> oldList, string attributeName, string attributewidth)
        {
            int outObj;
            string width = int.TryParse(attributewidth, out outObj) ? outObj.ToString() : "100";
            IEnumerable<savedqueryLayoutxmlGridRowCell> cells = oldList.Where(x => x.name.EndsWith(attributeName, StringComparison.InvariantCultureIgnoreCase));
            if (cells.Count() == 0)
            {
                cellList.Add(new savedqueryLayoutxmlGridRowCell()
                {
                    name = attributeName,
                    width = width
                });
            }
            else
            {
                savedqueryLayoutxmlGridRowCell currentcell = cells.First();
                cellList.Add(new savedqueryLayoutxmlGridRowCell()
                {
                    name = attributeName,
                    width = width,
                    LabelId = currentcell.LabelId,
                    label = currentcell.label,
                    ishidden = currentcell.ishidden,
                    disableSorting = currentcell.disableSorting,
                    disableMetaDataBinding = currentcell.disableMetaDataBinding,
                    desc = currentcell.desc,
                    cellType = currentcell.cellType,
                });
            }
        }


        public static CrmOperation generateUpdateRequest(string viewXml, string FetchXml, Guid viewId)
        {
            Entity sq = new Entity("savedquery");
            sq.Id = viewId;
            sq.Attributes.Add("fetchxml", FetchXml);
            sq.Attributes.Add("layoutxml", viewXml);
            UpdateRequest rq = new UpdateRequest();
            rq.Target = sq;
            return new CrmOperation(CrmOperation.CrmOperationType.update, CrmOperation.CrmOperationTarget.view, rq, "Update view");
        }

        public static CrmOperation generateUpdateRequest(string viewXml, string fetchXml, string name, string description, string logicalName)
        {
            Entity sq = new Entity("savedquery");
            sq.Attributes.Add("name", name);
            sq.Attributes.Add("description", description);
            sq.Attributes.Add("returnedtypecode", logicalName);
            sq.Attributes.Add("fetchxml", fetchXml);
            sq.Attributes.Add("layoutxml", viewXml);
            sq.Attributes.Add("querytype", 0);
            CreateRequest rq = new CreateRequest();
            rq.Target = sq;
            return new CrmOperation(CrmOperation.CrmOperationType.create, CrmOperation.CrmOperationTarget.view, rq, "Create new view");
        }


        public static string[] getEntityRelation(string cellData)
        {
            string relationtext = cellData;
            if (relationtext.LastIndexOf("-") > 0)
            {
                relationtext = relationtext.Substring(relationtext.LastIndexOf("-") + 1, relationtext.Length - relationtext.LastIndexOf("-") - 1);
                relationtext = relationtext.Trim();
                return relationtext.Split('.');
            }
            return null;
        }
        #endregion
    }
}
