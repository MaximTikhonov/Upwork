using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ContactPlugins
{
    public class FillFieldsOnCreate:  IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var executionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var crmService = factory.CreateOrganizationService(null);
            var target = (Entity)executionContext.InputParameters["Target"];
            if(!target.Contains("new_account"))
                return;
            var parentAccount = new Lazy<Entity>(() =>
            {
                 var accountId= target.GetAttributeValue<EntityReference>("new_account").Id;
                return crmService.Retrieve("account", accountId,
                    new ColumnSet(new[] {"emailaddress1", "telephone1", "name"}));
            });
            FillValues(target, parentAccount, "emailaddress1", "emailaddress1");
            FillValues(target, parentAccount, "telephone1", "telephone1");
        }

        private static void FillValues(Entity target, Lazy<Entity> parentEntity, string targetColumn, string sourceColumn)
        {
            if ((!target.Contains(targetColumn) || string.IsNullOrEmpty(target.GetAttributeValue<string>(targetColumn)))
                && !string.IsNullOrEmpty(parentEntity.Value.GetAttributeValue<string>(sourceColumn)))
            {
                target[targetColumn] = parentEntity.Value.GetAttributeValue<string>(sourceColumn);
            }
        }
    }
}
