using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace MergingAccountsActivity.WorkflowActivities
{
    public class LeadMergeActivity: CodeActivity
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
                new ConditionExpression("new_account", ConditionOperator.NotNull),
                new ConditionExpression("new_contact", ConditionOperator.Null),
                //new ConditionExpression("parentcustomerid", ConditionOperator.Equal,new Guid("F13486DB-711B-E511-80D0-3863BB353F28")),
                
            };
            var leads = crmConnector.GetLeads(new[] { "new_projectcontracttype", "new_projectdevelopmenttype", "new_projectgovermentreference", "new_contact", "new_account", "new_projectimportdate", "merged", "new_projectname", "new_projectno","subject","createdon" }, conditions);
            var leadsGrouped = leads.GroupBy(p => p.AccountId);
            //var leadsGrouped = leads.GroupBy(p => p.ContactId);
            var totalLeads = leads.LongCount();
            var errorText = new StringBuilder();
            int totalGroups = leadsGrouped.Count();
            int counter = 0;
            foreach (var iLeadGroup in leadsGrouped)
            {

                var leadsArray = iLeadGroup.OrderBy(p=>p.CreatedOn).ToArray();
                try
                {
                    MergingLogic.CalculateDuplicateLeads(crmConnector, leadsArray);
                }
                catch (Exception ex)
                {
                    errorText.AppendFormat("PostIndex {0}, Error: {1}", iLeadGroup.Key, ex.Message);
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
