using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class CrmOperation
    {
        public enum CrmOperationType { update, create, changeorder, error };
        public enum CrmOperationTarget { entity, attribute, optionSet,view , none };

        private CrmOperationType _operationType;
        private CrmOperationTarget _opertionTarget;
        private OrganizationRequest _orgRequest;
        private OrganizationResponse _orgResponse;
        private string _previewString;
        private string _exceptionString;
        private bool _executeOperation;



        public bool executeOperation
        {
            get { return _executeOperation; }
            set { _executeOperation = value; }
        }
        public string exceptionString
        {
            get { return _exceptionString; }
            set { _exceptionString = value; }
        }


        public string previewString
        {
            get { return _previewString; }
            set { _previewString = value; }
        }
        private bool _operationSucceded;

        public bool operationSucceded
        {
            get { return _operationSucceded; }
            set { _operationSucceded = value; }
        }

        public OrganizationRequest orgRequest
        {
            get { return _orgRequest; }
            set { _orgRequest = value; }
        }

        public OrganizationResponse orgResponse
        {
            get { return _orgResponse; }
            set { _orgResponse = value; }
        }


        public CrmOperation(CrmOperationType operationType, CrmOperationTarget opertionTarget, OrganizationRequest orgRequest, string previewString)
        {
            this.operationType = operationType;
            this.opertionTarget = opertionTarget;
            this.orgRequest = orgRequest;
            this.orgResponse = null;
            this.previewString = previewString;
            this.operationSucceded = false;
            this.executeOperation = true;

        }
        public CrmOperationTarget opertionTarget
        {
            get { return _opertionTarget; }
            set { _opertionTarget = value; }
        }

        public CrmOperationType operationType
        {
            get { return _operationType; }
            set { _operationType = value; }
        }

    }

    public class CrmUpdateAttributeOperation : CrmOperation
    {
        public CrmUpdateAttributeOperation(CrmOperationType operationType, CrmOperationTarget opertionTarget, IEnumerable<string> attributes, OrganizationRequest orgRequest, string previewString)
            : base(operationType, opertionTarget, orgRequest, previewString)
        {
            this.attributes = attributes;
        }

        private IEnumerable<string> _attributes;

        public IEnumerable<string> attributes
        {
            get { return _attributes; }
            set { _attributes = value; }
        }

    }


}
