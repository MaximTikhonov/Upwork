using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Common;
using MarketingListFactory.ProxyObjects.Excel;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Crm
{
    public class AccountCrmProxy: EntityCrmProxy, IAccountProxy
    {
        public AccountCrmProxy(Entity targetEntity) : base(targetEntity)
        {
        }

        public override bool IsEqual(ICompareEntity inputCompareEntity)
        {
            if (inputCompareEntity.GetType() != typeof (AccountCrmProxy))
                return false;
            var compareEntity = (AccountCrmProxy) inputCompareEntity;
            return (compareEntity.ShortName == this.ShortName)
                  || (!string.IsNullOrWhiteSpace(compareEntity.Email) && compareEntity.Email == this.Email)
                  || (!string.IsNullOrWhiteSpace(compareEntity.WebSite) && compareEntity.WebSite == this.WebSite)
                  || (!string.IsNullOrWhiteSpace(compareEntity.WorkPhone)
                      && (Extensions.PhoneCompare(this.WorkPhone, compareEntity.WorkPhone)
                          || Extensions.PhoneCompare(this.ExtraPhones, compareEntity.WorkPhone)))
                  || (!string.IsNullOrWhiteSpace(compareEntity.MobilePhone)
                      && (Extensions.PhoneCompare(this.MobilePhone, compareEntity.MobilePhone)
                          || Extensions.PhoneCompare(this.ExtraPhones, compareEntity.MobilePhone)))
                  || (compareEntity.PostCode == this.PostCode&& compareEntity.PostCode.Length > 1
                      && FuzzyStringComparer.IsStringsFuzzyEquals(compareEntity.ShortName, this.ShortName, 70));
        }

        public string Address1
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_line1"); }
            set { _targetEntity["address1_line1"] = value; }
        }

        public string Address2
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_line2"); }
            set { _targetEntity["address1_line2"] = value; }
        }
        public string Address3
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_line3"); }
            set { _targetEntity["address1_line3"] = value; }
        }
        public string CountyName
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_county"); }
            set { _targetEntity["address1_county"] = value; }
        }
        public string Email
        {
            get { return _targetEntity.GetAttributeValue<string>("emailaddress1"); }
            set { _targetEntity["emailaddress1"] = value; }
        }
        public string ExtraPhones
        {
            get { return _targetEntity.GetAttributeValue<string>("telephone3"); }
            set { _targetEntity["telephone3"] = value; }
        }
        public string ImportDate
        {
            get { return _targetEntity.GetAttributeValue<string>("new_excelimportdate"); }
            set { _targetEntity["new_excelimportdate"] = value; }
        }
        public Guid ImportId
        {
            get { return new Guid(_targetEntity.GetAttributeValue<string>("new_excelimportid")); }
            set { _targetEntity["new_excelimportid"] = value.ToString(); }
        }
        public string MobilePhone
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_telephone1"); }
            set { _targetEntity["address1_telephone1"] = value; }
        }
        public string Name
        {
            get { return _targetEntity.GetAttributeValue<string>("name"); }
            set { _targetEntity["name"] = value; }
        }
        public string PostCode
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_postalcode"); }
            set { _targetEntity["address1_postalcode"] = value; }
        }
        public string ShortName
        {
            get { return _targetEntity.GetAttributeValue<string>("new_shortname"); }
            set { _targetEntity["new_shortname"] = value; }
        }
        public string TownName
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_city"); }
            set { _targetEntity["address1_city"] = value; }
        }
        public string WebSite
        {
            get { return _targetEntity.GetAttributeValue<string>("websiteurl"); }
            set { _targetEntity["websiteurl"] = value; }
        }
        public string WorkPhone
        {
            get { return _targetEntity.GetAttributeValue<string>("telephone1"); }
            set { _targetEntity["telephone1"] = value; }
        }

        public static string[] ColumnsToArray()
        {
            return new[]
            {
                "telephone1", "websiteurl", "address1_city", "new_shortname", "address1_postalcode",
                "new_excelimportdate","new_excelimportid","address1_telephone1","name",
        "address1_line1","address1_line2","address1_line3","address1_county","emailaddress1","telephone3"
            };
        }

        public void Merge(IAccountProxy iAccountProxy)
        {
            
        }
    }
}
