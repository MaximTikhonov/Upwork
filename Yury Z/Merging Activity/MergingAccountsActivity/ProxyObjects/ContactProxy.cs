using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace MergingAccountsActivity.ProxyObjects
{
    public class ContactProxy : ICrmEntity
    {
        private readonly Entity _targetEntity;

        public ContactProxy(Entity targetEntity)
        {
            _targetEntity = targetEntity;
        }

        public Guid ParentCustomer
        {
            get { return _targetEntity.GetAttributeValue<EntityReference>("parentcustomerid").Id; }
        }

        public string CompanyPostCode
        {
            get
            {
                if(!_targetEntity.Contains("parentcustomer.address1_postalcode"))
                    return "";
                var val= _targetEntity.GetAttributeValue<AliasedValue>("parentcustomer.address1_postalcode").Value;
                if (val == null)
                    return "";
                return val.ToString().Replace(" ","");
            }
        }

        public bool IsEmptyParentAccount()
        {
            return string.IsNullOrWhiteSpace(_targetEntity.GetAttributeValue<EntityReference>("parentcustomerid").Name);
        }
        public string LastName
        {
            get
            {
                if (!_targetEntity.Contains("lastname"))
                    return "";
                return _targetEntity.GetAttributeValue<string>("lastname");
            }
        }
        public string FirstName
        {
            get
            {
                if (!_targetEntity.Contains("firstname"))
                    return "";
                return _targetEntity.GetAttributeValue<string>("firstname");
            }
            set { _targetEntity["firstname"] = value; }
        }
        public Guid Id {
            get { return _targetEntity.Id; }
        }
        public string Email {
            get
            {
                if (!_targetEntity.Contains("emailaddress1"))
                    return null;
                 return _targetEntity.GetAttributeValue<string>("emailaddress1");
            }
            set
            {
                _targetEntity["emailaddress1"] = value;
            }
        }
        public string Phone
        {
            get
            {
                if (!_targetEntity.Contains("telephone1"))
                    return null;
                return _targetEntity.GetAttributeValue<string>("telephone1");
            }
            set
            {
                _targetEntity["telephone1"] = value;
            }
        }
        public OptionSetValue Role
        {
            get
            {
                if (!_targetEntity.Contains("accountrolecode"))
                    return null;
                return _targetEntity.GetAttributeValue<OptionSetValue>("accountrolecode");
            }
            set
            {
                _targetEntity["accountrolecode"] = value;
            }
        }

        public DateTime CreatedOn
        {
            get { return _targetEntity.GetAttributeValue<DateTime>("createdon"); }
        }

        public Entity ToEntity()
        {
            return _targetEntity;
        }
    }
}
