using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MergingAccountsActivity.ProxyObjects;

namespace MergingAccountsActivity
{
    public static class MergingLogic
    {
        public static int EQUALITY_PERCENT;
        public static int MERGING_VERSION;

        public static void CalculateDuplicates(CrmConnector crmConnector, AccountProxy[] accountsArray)
        {
            var anotherNameAccounts = new List<AccountProxy>();
            var topAccount = accountsArray[0];
            for (int i = 1; i < accountsArray.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(accountsArray[i].ShortName))
                    continue;
                var computeResult = LevenshteinDistance.Compute(topAccount.ShortName, accountsArray[i].ShortName);
                if (1-(decimal)computeResult / accountsArray[i].ShortName.Length > (decimal)EQUALITY_PERCENT / 100)
                {
                    crmConnector.Merge(topAccount, accountsArray[i]);
                    crmConnector.SetMergedVersion(topAccount, MERGING_VERSION);
                    crmConnector.SetMergedVersion(accountsArray[i], MERGING_VERSION);
                }
                else
                {
                    anotherNameAccounts.Add(accountsArray[i]);
                }
            }
            if (anotherNameAccounts.Any())
            {
                var anotherNameAccountsArray = anotherNameAccounts.ToArray();
                CalculateDuplicates(crmConnector,  anotherNameAccountsArray);
            }
        }

        private static bool CompareFirstLetter(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first) && string.IsNullOrWhiteSpace(second))
                return true;
            if (string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second))
                return false;
            if (!string.IsNullOrWhiteSpace(first) && string.IsNullOrWhiteSpace(second))
                return false;
            if(first.Replace(" ","").ToLower()[0]==second.Replace(" ","").ToLower()[0])
                return true;
            return false;
        }
        public static void CalculateDuplicateContacts(CrmConnector crmConnector, ContactProxy[] contactsArray)
        {
            var anotherNameContacts = new List<ContactProxy>();
            var topContact = contactsArray[0];
            if (topContact.IsEmptyParentAccount())
                return;
            //if(topContact.Email.Contains("sent"))
                //return;
            for (int i = 1; i < contactsArray.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(contactsArray[i].LastName))
                    continue;
                var computeResult = LevenshteinDistance.Compute(topContact.LastName, contactsArray[i].LastName);
                if (1 - (decimal)computeResult / contactsArray[i].LastName.Length > (decimal)EQUALITY_PERCENT / 100
                   && CompareFirstLetter(topContact.FirstName, contactsArray[i].FirstName))
                {
                    crmConnector.MergeContact(topContact, contactsArray[i]);
                    crmConnector.SetMergedVersion(topContact, MERGING_VERSION);
                    crmConnector.SetMergedVersion(contactsArray[i], MERGING_VERSION);
                }
                else
                {
                    anotherNameContacts.Add(contactsArray[i]);
                }
            }
            if (anotherNameContacts.Any())
            {
                var anotherNameContactsArray = anotherNameContacts.ToArray();
                CalculateDuplicateContacts(crmConnector, anotherNameContactsArray);
            }
        }

        public static void CalculateDuplicateLeads(CrmConnector crmConnector, LeadProxy[] leadsArray)
        {
            var anotherNameLeads = new List<LeadProxy>();
            var topLead = leadsArray[0];
            if(string.IsNullOrWhiteSpace(topLead.Topic))
                return;
            for (int i = 1; i < leadsArray.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(leadsArray[i].Topic))
                    continue;
                var computeResult = LevenshteinDistance.Compute(topLead.Topic, leadsArray[i].Topic);
                if (1 - (decimal)computeResult / leadsArray[i].Topic.Length > (decimal)EQUALITY_PERCENT / 100)
                {
                    crmConnector.MergeLead(topLead, leadsArray[i]);
                    crmConnector.SetMergedVersion(topLead, MERGING_VERSION);
                    crmConnector.SetMergedVersion(leadsArray[i], MERGING_VERSION);
                }
                else
                {
                    anotherNameLeads.Add(leadsArray[i]);
                }
            }
            if (anotherNameLeads.Any())
            {
                var anotherSubjectLeadArray = anotherNameLeads.ToArray();
                CalculateDuplicateLeads(crmConnector, anotherSubjectLeadArray);
            }
        }
    }
}
