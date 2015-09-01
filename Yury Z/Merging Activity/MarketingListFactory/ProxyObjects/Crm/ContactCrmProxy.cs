using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Resources;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Common;
using MarketingListFactory.ProxyObjects.Excel;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Crm
{
    public class ContactCrmProxy:EntityCrmProxy, IContactProxy
    {
        private decimal EQUALITY_PERCENT=80;

        public ContactCrmProxy(Entity targetEntity)
            : base(targetEntity)
        {
        }

        public override bool IsEqual(ICompareEntity inputCompareEntity)
        {
            if (inputCompareEntity.GetType() != typeof(ContactCrmProxy))
                return false;
            var compareEntity = (ContactCrmProxy)inputCompareEntity;
            if(string.IsNullOrEmpty(this.LastName)||string.IsNullOrEmpty(compareEntity.LastName))
                return false;
            var computeResult = LevenshteinDistance.Compute(this.LastName, compareEntity.LastName);
            if (1 - (decimal)computeResult / compareEntity.LastName.Length > (decimal)EQUALITY_PERCENT / 100
               && CompareFirstLetter(this.FirstName, compareEntity.FirstName))
            {
                return true;
            }
            return false;
        }

        public string FirstName
        {
            get { return _targetEntity.GetAttributeValue<string>("firstname"); }
            set { _targetEntity["firstname"] = value; }
        }

        public string LastName
        {
            get { return _targetEntity.GetAttributeValue<string>("lastname"); }
            set { _targetEntity["lastname"] = value; }
        }

        private static bool CompareFirstLetter(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first) && string.IsNullOrWhiteSpace(second))
                return true;
            if (string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second))
                return false;
            if (!string.IsNullOrWhiteSpace(first) && string.IsNullOrWhiteSpace(second))
                return false;
            if (first.Replace(" ", "").ToLower()[0] == second.Replace(" ", "").ToLower()[0])
                return true;
            return false;
        }
        public static string[] ColumnsToArray()
        {
            return new[]
            {
                "new_excelimportid",
                "telephone1",
                "new_shortfirstname",
                "business2",
                "fax",
                "websiteurl",
                "firstname",
                "emailaddress1",
                "suffix",
                "lastname",
                "parentcustomerid",
                "accountrolecode"
            };
        }

        public void Merge(ContactExcelProxy iContactProxy)
        {
            //throw new System.NotImplementedException();
        }
    }
}
