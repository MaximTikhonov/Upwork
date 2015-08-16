using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace MergingAccountsActivity.WorkflowActivities
{
    public class ContactMergeActivity: CodeActivity
    {
        [Input("RecursiveJob")]
        [ReferenceTarget("new_recursivejob")]
        [RequiredArgument]
        public InArgument<EntityReference> RecursiveJob { get; set; }
        [Input("EqualityPercentage")]
        [RequiredArgument]
        public InArgument<int> EqualityPercentage { get; set; }
        [Input("CrmService")]
        public InArgument<IOrganizationService> CrmService { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            var equalityPercentage = EqualityPercentage.Get(executionContext);
            MergingLogic.EQUALITY_PERCENT = equalityPercentage;
            var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            //var crmService = serviceFactory.CreateOrganizationService(Guid.Empty);
            var crmService = CrmService.Get(executionContext);
            var crmConnector = new CrmConnector(crmService);
            var recursiveEntity = crmConnector.GetRecursiveEntity(RecursiveJob.Get(executionContext), new[] { "new_lastlaunchdate", "new_launchindex" });
            var mergingVersion = recursiveEntity.GetAttributeValue<int>("new_launchindex");
            MergingLogic.MERGING_VERSION = mergingVersion;
            var conditions = new List<ConditionExpression>()
            {
                new ConditionExpression("merged", ConditionOperator.NotEqual,true),
                //new ConditionExpression("statecode", ConditionOperator.Equal,0),
                //new ConditionExpression("new_lastmergeindex", ConditionOperator.Null),
                new ConditionExpression("parentcustomerid", ConditionOperator.NotNull),
                //new ConditionExpression("emailaddress1", ConditionOperator.NotNull),
                //new ConditionExpression("parentcustomerid", ConditionOperator.Equal,new Guid("F13486DB-711B-E511-80D0-3863BB353F28")),
                
            };
            var contacts =
                crmConnector.GetContacts(
                    new[]
                    {
                        "lastname", "firstname", "parentcustomerid", "telephone1", "emailaddress1", "accountrolecode",
                        "merged", "createdon"
                    }, conditions);
            var contactsByCompany = contacts.GroupBy(p => p.ParentCustomer);
            //var contactsByCompany = contacts.GroupBy(p => p.Email);
            var totalContacts = contacts.LongCount();
            var errorText = new StringBuilder();
            int totalGroups = contactsByCompany.Count();
            int counter = 0;
            foreach (var iContactGroup in contactsByCompany)
            {
                
                var contactsArray = iContactGroup.OrderBy(p=>p.CreatedOn).ToArray();
                try
                {
                    MergingLogic.CalculateDuplicateContacts(crmConnector, contactsArray);
                }
                catch (Exception ex)
                {
                    errorText.AppendFormat("PostIndex {0}, Error: {1}", iContactGroup.Key, ex.Message);
                    errorText.AppendLine();
                }
                counter++;
            }
            recursiveEntity["new_launchindex"] = mergingVersion + 1;
            recursiveEntity["new_lastlaunchdate"] = DateTime.Now;
            recursiveEntity["new_errortext"] = errorText.ToString();
            //crmConnector.Update(recursiveEntity);
        }
    }
}
