using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Channels;
using System.Text;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Common;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Crm
{
    public class LeadCrmProxy : EntityCrmProxy, IProject
    {
        private decimal EQUALITY_PERCENT=70;

        public LeadCrmProxy(Entity targetLead)
            : base(targetLead)
        {

        }

        public override bool IsEqual(ICompareEntity compareEntity)
        {
            if(compareEntity.GetType()!=typeof(LeadCrmProxy))
                return false;
            var compareLead = (LeadCrmProxy) compareEntity;
            if(string.IsNullOrEmpty(this.Name)||string.IsNullOrEmpty(compareLead.Name))
                return false;
             var computeResult = LevenshteinDistance.Compute(this.Name, compareLead.Name);
            if (1 - (decimal) computeResult/compareLead.Name.Length > (decimal) EQUALITY_PERCENT/100)
            {
                return true;
            }
            return false;
        }

        public static string[] ColumnsToArray()
        {
            return new[]
            {
                "new_excelimportid",
                "subject",
                "new_topic2",
                "new_projectimportdate",
                "description",
                "emailaddress1",
                "firstname",
                "telephone1",
                "lastname",
                "new_account",
                "new_contact"
            };
        }

        public string Email
        {
            get { return _targetEntity.GetAttributeValue<string>("emailaddress1"); }
            set { _targetEntity["emailaddress1"] = value; }
        }

        public string Name
        {
            get { return _targetEntity.GetAttributeValue<string>("subject"); }
            set { _targetEntity["subject"] = value; }
        }
    }
}
