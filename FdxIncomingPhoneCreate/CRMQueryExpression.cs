using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxIncomingPhoneCreate
{
    class CRMQueryExpression
    {
        string filterAttribute;
        ConditionOperator filterOperator;
        object filterValue;

        public CRMQueryExpression(string _attr, ConditionOperator _opr, object _val)
        {
            this.filterAttribute = _attr;
            this.filterOperator = _opr;
            this.filterValue = _val;
        }
        public static QueryExpression getQueryExpression(string _entityName, ColumnSet _columnSet, CRMQueryExpression[] _exp)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = _columnSet;
            query.Distinct = false;
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;
            for (int i = 0; i < _exp.Length; i++)
            {
                query.Criteria.AddCondition(_exp[i].filterAttribute, _exp[i].filterOperator, _exp[i].filterValue);
            }

            return query;
        }

        public static int GetOptionsetValue(IOrganizationService _service, string _optionsetName, string _optionsetSelectedText)
        {
            int optionsetValue = 0;
            try
            {
                RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = _optionsetName
                    };

                // Execute the request.
                RetrieveOptionSetResponse retrieveOptionSetResponse =
                    (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    if (optionMetadata.Label.UserLocalizedLabel.Label.ToString() == _optionsetSelectedText)
                    {
                        optionsetValue = (int)optionMetadata.Value;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return optionsetValue;
        }
        public static string GetOptionsSetTextForValue(IOrganizationService service, string entityName, string attributeName, int selectedValue)
        {

            RetrieveAttributeRequest retrieveAttributeRequest = new
            RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            // Access the retrieved attribute.
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)
            retrieveAttributeResponse.AttributeMetadata;// Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
            string selectedOptionLabel = null;
            foreach (OptionMetadata oMD in optionList)
            {
                if (oMD.Value == selectedValue)
                {
                    selectedOptionLabel = oMD.Label.LocalizedLabels[0].Label.ToString();
                    break;
                }
            }
            return selectedOptionLabel;
        }
    }
}
