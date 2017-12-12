using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxIncomingPhoneCreate
{
    public class IncomingPhoneCampaignResponseCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            int step = 1;
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    step = 2;
                    Entity phoneCall = (Entity)context.InputParameters["Target"];
                    //Check if the Entity on which plugin is executing is Phone Call
                    if (phoneCall.LogicalName != "phonecall")
                        return;
                    //Check if the Phone Call record is Completed either with Contacted or Demo Set status
                    //if (phoneCall.Attributes.Contains("statuscode") && (((OptionSetValue)phoneCall.Attributes["statuscode"]).Value == 2 || ((OptionSetValue)phoneCall.Attributes["statuscode"]).Value == 756480004 || ((OptionSetValue)phoneCall.Attributes["statuscode"]).Value == 756480005))
                    if (phoneCall.Attributes.Contains("category") && (phoneCall.Attributes["category"].ToString() == "Contacted" || phoneCall.Attributes["category"].ToString() == "Demo Set" || phoneCall.Attributes["category"].ToString() == "Transferred")) 
                    {
                        step = 3;

                        //Get the phone call details ( All Attributes )
                        QueryExpression phoneCallQuery = new QueryExpression();
                        phoneCallQuery.EntityName = phoneCall.LogicalName;
                        phoneCallQuery.ColumnSet = new ColumnSet(true);
                        phoneCallQuery.Criteria.AddCondition("activityid", ConditionOperator.Equal, (Guid)phoneCall.Id);

                        Entity phoneCallRecord = service.RetrieveMultiple(phoneCallQuery).Entities[0];

                        //Go into the loop only if phone call is of type Incoming (false)
                        if (phoneCallRecord.Attributes.Contains("directioncode") && (bool)phoneCallRecord.Attributes["directioncode"] == false)
                        {
                            //Go into the loop only if phonecall contains dnis
                            if (phoneCallRecord.Attributes.Contains("fdx_dnis") && phoneCallRecord.Attributes["fdx_dnis"] != null)
                            {
                                step = 4;
                                //Query to pick recent campaign with the DNIS mentioned in phonecall
                                QueryExpression campaignQuery = new QueryExpression();
                                campaignQuery.EntityName = "campaign";
                                campaignQuery.ColumnSet = new ColumnSet("campaignid", "codename", "name");
                                campaignQuery.Criteria.AddFilter(LogicalOperator.And);
                                campaignQuery.Criteria.AddCondition("fdx_dnis", ConditionOperator.Equal, phoneCallRecord.Attributes["fdx_dnis"].ToString());
                                campaignQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                campaignQuery.AddOrder("createdon", OrderType.Descending);

                                EntityCollection campaignRecordSet = service.RetrieveMultiple(campaignQuery);

                                //Check if there exist a campaign with the dnis mentioned in Phonecall
                                if (campaignRecordSet.Entities.Count > 0)
                                {
                                    Entity campaignRecord = campaignRecordSet[0];

                                    //Check if phone call is tagged with a Record Type (any entity Type will be tagged in the same field. But current requirement is lead/opp/contact)
                                    if (phoneCallRecord.Attributes.Contains("regardingobjectid"))
                                    {
                                        step = 5;
                                        //Create a Campaign Response object, which needs to be created as part of Incoming Phone Call against a DNIS

                                        Entity campaignResponse = new Entity("campaignresponse");
                                        campaignResponse.Attributes["channeltypecode"] = new OptionSetValue(3);//Type of Campaign Response : Phone Call(3)
                                        campaignResponse.Attributes["regardingobjectid"] = new EntityReference(campaignRecord.LogicalName, campaignRecord.Id);
                                        campaignResponse.Attributes["fdx_parentcampaigncode"] = campaignRecord.Attributes["codename"].ToString();
                                        campaignResponse.Attributes["receivedon"] = DateTime.UtcNow;
                                        string switchEntity = ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).LogicalName;

                                        step = 6;
                                        //Set attribute values in Campaign Response object based on the regardignobject type in Phone Call
                                        switch (switchEntity)
                                        {
                                            case "lead":
                                            
                                                #region Create a CR by querying and copying details from a Lead
                                                
                                                step = 7;
                                                //Query the Lead and get the required attributes to set into Campaign Response object
                                                QueryExpression queryLead = new QueryExpression();
                                                queryLead.EntityName = switchEntity;
                                                queryLead.ColumnSet = new ColumnSet("firstname", "lastname", "fdx_credential", "telephone1", "telephone2", "fdx_jobtitlerole", "emailaddress1", "address1_line1", "address1_line2", "address1_city", "fdx_stateprovince", "fdx_zippostalcode", "address1_country", "telephone3", "companyname", "websiteurl", "campaignid","leadsourcecode");
                                                queryLead.Criteria.AddCondition("leadid", ConditionOperator.Equal, ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).Id);

                                                step = 8;
                                                Entity lead = new Entity();
                                                lead = service.RetrieveMultiple(queryLead).Entities[0];
                                                step = 9;
                                                string firstname = "";
                                                string lastname = "";
                                                if (lead.Attributes.Contains("firstname"))
                                                {
                                                    firstname = lead.Attributes["firstname"].ToString();
                                                    campaignResponse.Attributes["firstname"] = lead.Attributes["firstname"].ToString();
                                                }
                                                if (lead.Attributes.Contains("lastname"))
                                                {
                                                    lastname = lead.Attributes["lastname"].ToString();
                                                    campaignResponse.Attributes["lastname"] = lead.Attributes["lastname"].ToString();
                                                }
                                                campaignResponse.Attributes["subject"] = firstname + " " + lastname;

                                                step = 10;
                                                campaignResponse.Attributes["fdx_reconversionlead"] = new EntityReference(lead.LogicalName, lead.Id);
                                                if (lead.Attributes.Contains("fdx_credential"))
                                                {
                                                    campaignResponse.Attributes["fdx_credential"] = lead.Attributes["fdx_credential"].ToString();
                                                }
                                                if (lead.Attributes.Contains("telephone1"))
                                                {
                                                    campaignResponse.Attributes["fdx_telephone1"] = lead.Attributes["telephone1"].ToString();
                                                }
                                                step = 11;
                                                if (lead.Attributes.Contains("fdx_jobtitlerole"))
                                                {
                                                    campaignResponse.Attributes["fdx_jobtitlerole"] = new OptionSetValue(((OptionSetValue)(lead.Attributes["fdx_jobtitlerole"])).Value);
                                                }
                                                if (lead.Attributes.Contains("telephone2"))
                                                {
                                                    campaignResponse.Attributes["telephone"] = lead.Attributes["telephone2"].ToString();
                                                }
                                                if (lead.Attributes.Contains("emailaddress1"))
                                                {
                                                    campaignResponse.Attributes["emailaddress"] = lead.Attributes["emailaddress1"].ToString();
                                                }
                                                if (lead.Attributes.Contains("companyname"))
                                                {
                                                    campaignResponse.Attributes["companyname"] = lead.Attributes["companyname"].ToString();
                                                }
                                                step = 12;
                                                if (lead.Attributes.Contains("fdx_zippostalcode"))
                                                {
                                                    campaignResponse.Attributes["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", ((EntityReference)lead.Attributes["fdx_zippostalcode"]).Id);
                                                }
                                                step = 13;
                                                if (lead.Attributes.Contains("websiteurl"))
                                                {
                                                    campaignResponse.Attributes["fdx_websiteurl"] = lead.Attributes["websiteurl"].ToString();
                                                }
                                                if (lead.Attributes.Contains("telephone3"))
                                                {
                                                    campaignResponse.Attributes["fdx_telephone3"] = lead.Attributes["telephone3"].ToString();
                                                }
                                                if (lead.Attributes.Contains("address1_line1"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line1"] = lead.Attributes["address1_line1"].ToString();
                                                }
                                                if (lead.Attributes.Contains("address1_line2"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line2"] = lead.Attributes["address1_line2"].ToString();
                                                }
                                                if (lead.Attributes.Contains("address1_city"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_city"] = lead.Attributes["address1_city"].ToString();
                                                }
                                                step = 14;
                                                if (lead.Attributes.Contains("fdx_stateprovince"))
                                                {
                                                    campaignResponse.Attributes["fdx_stateprovince"] = new EntityReference("fdx_state", ((EntityReference)lead.Attributes["fdx_stateprovince"]).Id);
                                                }
                                                step = 15;
                                                if (lead.Attributes.Contains("address1_country"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_country"] = lead.Attributes["address1_country"].ToString();
                                                }
                                                step = 16;
                                                service.Create(campaignResponse);
                                                step = 17;

                                                #endregion

                                                #region Update Lead's source campaign for first campaign response

                                                if (!lead.Attributes.Contains("campaignid") && ((OptionSetValue)lead.Attributes["leadsourcecode"]).Value == 3)
                                                {
                                                    step = 171;
                                                    Entity LeadUpdate = new Entity("lead");

                                                    LeadUpdate.Id = ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).Id;
                                                    LeadUpdate.Attributes["campaignid"] = new EntityReference("campaign", ((EntityReference)campaignResponse.Attributes["regardingobjectid"]).Id);
                                                    step = 172;
                                                    service.Update(LeadUpdate);
                                                }
                                               
                                                #endregion
                                                
                                                break;
                                            case "contact":

                                            #region Create a CR by querying and copying details from a Contact
                                                
                                                step = 18;
                                                //Query the Contact and get the required attributes to set into Campaign Response object
                                                QueryExpression queryContact = new QueryExpression();
                                                queryContact.EntityName = switchEntity;
                                                queryContact.ColumnSet = new ColumnSet("firstname", "lastname", "fdx_credential", "telephone1", "telephone2", "fdx_jobtitlerole", "emailaddress1", "address1_line1", "address1_line2", "address1_city", "fdx_stateprovinceid", "fdx_zippostalcodeid", "address1_country", "websiteurl");
                                                queryContact.Criteria.AddCondition("contactid", ConditionOperator.Equal, ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).Id);

                                                step = 19;
                                                Entity contact = new Entity();
                                                contact = service.RetrieveMultiple(queryContact).Entities[0];
                                                step = 20;
                                                firstname = "";
                                                lastname = "";
                                                if (contact.Attributes.Contains("firstname"))
                                                {
                                                    firstname = contact.Attributes["firstname"].ToString();
                                                    campaignResponse.Attributes["firstname"] = contact.Attributes["firstname"].ToString();
                                                }
                                                if (contact.Attributes.Contains("lastname"))
                                                {
                                                    lastname = contact.Attributes["lastname"].ToString();
                                                    campaignResponse.Attributes["lastname"] = contact.Attributes["lastname"].ToString();
                                                }
                                                campaignResponse.Attributes["subject"] = firstname + " " + lastname;

                                                step = 21;
                                                campaignResponse.Attributes["fdx_reconversioncontact"] = new EntityReference(contact.LogicalName, contact.Id);
                                                if (contact.Attributes.Contains("fdx_credential"))
                                                {
                                                    campaignResponse.Attributes["fdx_credential"] = contact.Attributes["fdx_credential"].ToString();
                                                }
                                                if (contact.Attributes.Contains("telephone1"))
                                                {
                                                    campaignResponse.Attributes["fdx_telephone1"] = contact.Attributes["telephone1"].ToString();
                                                }
                                                step = 22;
                                                if (contact.Attributes.Contains("fdx_jobtitlerole"))
                                                {
                                                    campaignResponse.Attributes["fdx_jobtitlerole"] = new OptionSetValue(((OptionSetValue)(contact.Attributes["fdx_jobtitlerole"])).Value);
                                                }
                                                if (contact.Attributes.Contains("telephone2"))
                                                {
                                                    campaignResponse.Attributes["telephone"] = contact.Attributes["telephone2"].ToString();
                                                }
                                                if (contact.Attributes.Contains("emailaddress1"))
                                                {
                                                    campaignResponse.Attributes["emailaddress"] = contact.Attributes["emailaddress1"].ToString();
                                                }
                                                step = 23;
                                                if (contact.Attributes.Contains("fdx_zippostalcode"))
                                                {
                                                    campaignResponse.Attributes["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", ((EntityReference)contact.Attributes["fdx_zippostalcodeid"]).Id);
                                                }
                                                step = 24;
                                                if (contact.Attributes.Contains("websiteurl"))
                                                {
                                                    campaignResponse.Attributes["fdx_websiteurl"] = contact.Attributes["websiteurl"].ToString();
                                                }
                                                if (contact.Attributes.Contains("address1_line1"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line1"] = contact.Attributes["address1_line1"].ToString();
                                                }
                                                if (contact.Attributes.Contains("address1_line2"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line2"] = contact.Attributes["address1_line2"].ToString();
                                                }
                                                if (contact.Attributes.Contains("address1_city"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_city"] = contact.Attributes["address1_city"].ToString();
                                                }
                                                step = 25;
                                                if (contact.Attributes.Contains("fdx_stateprovinceid"))
                                                {
                                                    campaignResponse.Attributes["fdx_stateprovince"] = new EntityReference("fdx_state", ((EntityReference)contact.Attributes["fdx_stateprovinceid"]).Id);
                                                }
                                                step = 26;
                                                if (contact.Attributes.Contains("address1_country"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_country"] = contact.Attributes["address1_country"].ToString();
                                                }
                                                step = 27;
                                                service.Create(campaignResponse);
                                            
                                                #endregion
                                            
                                                step = 28;
                                                break;
                                            case "opportunity":
                                                
                                                #region Create a CR by querying and copying details from an Opportunity
                                                
                                                step = 29;
                                                //Query the Opportunity and get the Account and Contact of Opportunity to set into Campaign Response object
                                                QueryExpression queryOpportunity = new QueryExpression();
                                                queryOpportunity.EntityName = switchEntity;
                                                queryOpportunity.ColumnSet = new ColumnSet("parentaccountid", "parentcontactid");
                                                queryOpportunity.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).Id);

                                                step = 30;
                                                Entity opportunity = new Entity();
                                                opportunity = service.RetrieveMultiple(queryOpportunity).Entities[0];
                                                step = 31;
                                                Entity opportunityContact = new Entity();
                                                Entity opportunityAccount = new Entity();

                                                //Query the Account related to Opportunity to get the Address related attributes
                                                if (opportunity.Attributes.Contains("parentaccountid"))
                                                {
                                                    step = 32;
                                                    QueryExpression queryOpportunityAccount = new QueryExpression();
                                                    queryOpportunityAccount.EntityName = "account";
                                                    queryOpportunityAccount.ColumnSet = new ColumnSet("address1_line1", "address1_line2", "address1_city", "fdx_stateprovinceid", "fdx_zippostalcodeid", "address1_country", "websiteurl");

                                                    queryOpportunityAccount.Criteria.AddCondition("accountid", ConditionOperator.Equal, ((EntityReference)opportunity.Attributes["parentaccountid"]).Id);
                                                    step = 33;
                                                    opportunityAccount = service.RetrieveMultiple(queryOpportunityAccount).Entities[0];
                                                    step = 34;
                                                }

                                                //Query the Contact related to Opportunity to get the Phone or other attributes
                                                if (opportunity.Attributes.Contains("parentcontactid"))
                                                {
                                                    step = 35;
                                                    QueryExpression queryOpportunityContact = new QueryExpression();
                                                    queryOpportunityContact.EntityName = "contact";
                                                    queryOpportunityContact.ColumnSet = new ColumnSet("firstname", "lastname", "fdx_credential", "telephone1", "telephone2", "fdx_jobtitlerole", "emailaddress1");
                                                    queryOpportunityContact.Criteria.AddCondition("contactid", ConditionOperator.Equal, ((EntityReference)opportunity.Attributes["parentcontactid"]).Id);
                                                    step = 36;
                                                    opportunityContact = service.RetrieveMultiple(queryOpportunityContact).Entities[0];
                                                    step = 37;
                                                }

                                                firstname = "";
                                                lastname = "";
                                                if (opportunityContact.Attributes.Contains("firstname"))
                                                {
                                                    firstname = opportunityContact.Attributes["firstname"].ToString();
                                                    campaignResponse.Attributes["firstname"] = opportunityContact.Attributes["firstname"].ToString();
                                                    step = 38;
                                                }
                                                if (opportunityContact.Attributes.Contains("lastname"))
                                                {
                                                    lastname = opportunityContact.Attributes["lastname"].ToString();
                                                    campaignResponse.Attributes["lastname"] = opportunityContact.Attributes["lastname"].ToString();
                                                    step = 39;
                                                }
                                                campaignResponse.Attributes["subject"] = firstname + " " + lastname;

                                                step = 40;
                                                campaignResponse.Attributes["fdx_reconversionopportunity"] = new EntityReference(opportunity.LogicalName, opportunity.Id);
                                                if (opportunityContact.Attributes.Contains("fdx_credential"))
                                                {
                                                    campaignResponse.Attributes["fdx_credential"] = opportunityContact.Attributes["fdx_credential"].ToString();
                                                }
                                                if (opportunityContact.Attributes.Contains("telephone1"))
                                                {
                                                    campaignResponse.Attributes["fdx_telephone1"] = opportunityContact.Attributes["telephone1"].ToString();
                                                }
                                                step = 41;
                                                if (opportunityContact.Attributes.Contains("fdx_jobtitlerole"))
                                                {
                                                    campaignResponse.Attributes["fdx_jobtitlerole"] = new OptionSetValue(((OptionSetValue)(opportunityContact.Attributes["fdx_jobtitlerole"])).Value);
                                                }
                                                if (opportunityContact.Attributes.Contains("telephone2"))
                                                {
                                                    campaignResponse.Attributes["telephone"] = opportunityContact.Attributes["telephone2"].ToString();
                                                }
                                                if (opportunityContact.Attributes.Contains("emailaddress1"))
                                                {
                                                    campaignResponse.Attributes["emailaddress"] = opportunityContact.Attributes["emailaddress1"].ToString();
                                                }
                                                step = 42;
                                                if (opportunityAccount.Attributes.Contains("fdx_zippostalcode"))
                                                {
                                                    campaignResponse.Attributes["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", ((EntityReference)opportunityAccount.Attributes["fdx_zippostalcodeid"]).Id);
                                                }
                                                step = 43;
                                                if (opportunityAccount.Attributes.Contains("websiteurl"))
                                                {
                                                    campaignResponse.Attributes["fdx_websiteurl"] = opportunityAccount.Attributes["websiteurl"].ToString();
                                                }
                                                if (opportunityAccount.Attributes.Contains("address1_line1"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line1"] = opportunityAccount.Attributes["address1_line1"].ToString();
                                                }
                                                if (opportunityAccount.Attributes.Contains("address1_line2"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_line2"] = opportunityAccount.Attributes["address1_line2"].ToString();
                                                }
                                                if (opportunityAccount.Attributes.Contains("address1_city"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_city"] = opportunityAccount.Attributes["address1_city"].ToString();
                                                }
                                                step = 44;
                                                if (opportunityAccount.Attributes.Contains("fdx_stateprovinceid"))
                                                {
                                                    campaignResponse.Attributes["fdx_stateprovince"] = new EntityReference("fdx_state", ((EntityReference)opportunityAccount.Attributes["fdx_stateprovinceid"]).Id);
                                                }
                                                step = 45;
                                                if (opportunityAccount.Attributes.Contains("address1_country"))
                                                {
                                                    campaignResponse.Attributes["fdx_address1_country"] = opportunityAccount.Attributes["address1_country"].ToString();
                                                }
                                                step = 46;

                                                service.Create(campaignResponse);
                                                
                                                #endregion

                                                //#region Update Opportunity's source campaign for first campaign response
                                                //if (!opportunity.Attributes.Contains("campaignid"))
                                                //{
                                                //    step = 47;
                                                //    Entity OpportunityUpdate = new Entity("opportunity");

                                                //    OpportunityUpdate.Id = ((EntityReference)phoneCallRecord.Attributes["regardingobjectid"]).Id;
                                                //    OpportunityUpdate.Attributes["campaignid"] = new EntityReference("campaign", ((EntityReference)campaignResponse.Attributes["regardingobjectid"]).Id);
                                                //    step = 48;
                                                //    service.Update(OpportunityUpdate);
                                                //}
                                                //#endregion


                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(string.Format("An error occurred in the IncomingPhoneCampaignResponseCreate plug-in at Step {0}.", step), ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("IncomingPhoneCampaignResponseCreate: step {0}, {1}", step, ex.ToString());
                throw;
            }
        }
    }
}
