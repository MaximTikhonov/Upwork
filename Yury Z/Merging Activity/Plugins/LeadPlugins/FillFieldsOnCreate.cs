using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace LeadPlugins
{
    public class FillFieldsOnCreate: IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var executionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var crmService = factory.CreateOrganizationService(null);
            var target = (Entity)executionContext.InputParameters["Target"];
            var parentAccount = new Lazy<Entity>(() =>
            {
                if (!target.Contains("new_account"))
                    return new Entity("account");
                var accountId = target.GetAttributeValue<EntityReference>("new_account").Id;
                return crmService.Retrieve("account", accountId,
                    new ColumnSet(new[] { "emailaddress1", "telephone1", "name" }));
            });
            var parentContact = new Lazy<Entity>(() =>
            {
                if (!target.Contains("new_contact"))
                    return new Entity("contact");
                var accountId = target.GetAttributeValue<EntityReference>("new_contact").Id;
                return crmService.Retrieve("contact", accountId,
                    new ColumnSet(new[] { "emailaddress1", "telephone1", "firstname","lastname" }));
            });
            FillValues(target, parentContact, "emailaddress1", "emailaddress1");
            FillValues(target, parentContact, "telephone1", "telephone1");
            FillValues(target, parentAccount, "emailaddress1", "emailaddress1");
            FillValues(target, parentAccount, "telephone1", "telephone1");
            FillValues(target, parentAccount, "companyname", "name");
            FillValues(target, parentContact, "firstname", "firstname");
            FillValues(target, parentContact, "lastname", "lastname");
        }
        private static void FillValues(Entity target, Lazy<Entity> parentEntity, string targetColumn, string sourceColumn)
        {
            if ((!target.Contains(targetColumn)||string.IsNullOrEmpty(target.GetAttributeValue<string>(targetColumn)))
                && !string.IsNullOrEmpty(parentEntity.Value.GetAttributeValue<string>(sourceColumn)))
            {
                target[targetColumn] = parentEntity.Value.GetAttributeValue<string>(sourceColumn);
            }
        }
    }
}
