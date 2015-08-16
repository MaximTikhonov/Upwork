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
    public class MergeAccountsActivity: CodeActivity
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
            var crmConnector=new CrmConnector(crmService);
            var recursiveEntity = crmConnector.GetRecursiveEntity(RecursiveJob.Get(executionContext), new[] { "new_lastlaunchdate", "new_launchindex" });
            var mergingVersion = recursiveEntity.GetAttributeValue<int>("new_launchindex");
            MergingLogic.MERGING_VERSION = mergingVersion;
            var conditions=new List<ConditionExpression>()
            {
                new ConditionExpression("merged", ConditionOperator.NotEqual,true),
                //new ConditionExpression("new_lastmergeindex", ConditionOperator.NotEqual, mergingVersion),
                new ConditionExpression("address1_postalcode", ConditionOperator.NotNull),
            };
            var accounts = crmConnector.GetAccounts(new[] { "address1_postalcode", "new_shortname", "createdon" }, conditions);
            var accountsByPostCode = accounts.GroupBy(p => p.PostCode);
            var errorText=new StringBuilder();
            int totalGroups = accountsByPostCode.Count();
            int counter = 0;
            foreach (var iAccountGroup in accountsByPostCode)
            {
                var accountsArray = iAccountGroup.OrderBy(p=>p.CreatedOn).ToArray();
                try
                {
                    MergingLogic.CalculateDuplicates(crmConnector, accountsArray);
                }
                catch (Exception ex)
                {
                    errorText.AppendFormat("PostIndex {0}, Error: {1}",iAccountGroup.Key,ex.Message);
                    errorText.AppendLine();
                }
                counter++;
            }
            recursiveEntity["new_launchindex"] = mergingVersion+1;
            recursiveEntity["new_lastlaunchdate"] =DateTime.Now;
            recursiveEntity["new_errortext"] = errorText.ToString();
            crmConnector.Update(recursiveEntity);
        }
    }
}
