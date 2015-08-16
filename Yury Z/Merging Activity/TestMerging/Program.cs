using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MergingAccountsActivity;
using MergingAccountsActivity.WorkflowActivities;
using Microsoft.Xrm.Sdk;

namespace TestMerging
{
    class Program
    {
        static void Main(string[] args)
        {
            var service = CrmService.CreateOrganizationService();
            var imput = new Dictionary<string, object>()
            {
                {"RecursiveJob", new EntityReference("new_recursivejob",new Guid("30D602D7-F220-E511-80D0-3863BB34E990"))},
                {"EqualityPercentage",80},
                {"CrmService", service}
            };
            //var act = new MergeAccountsActivity();
            //var act = new ContactMergeActivity();
            var act=new LeadMergeActivity();
            WorkflowInvoker.Invoke(act, imput);
            
            
        }
    }
}
