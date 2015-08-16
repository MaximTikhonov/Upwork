using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory
{
    public class LeadProxy
    {
        private readonly Entity _targetEntity;

        public LeadProxy(Entity targetEntity)
        {
            _targetEntity = targetEntity;
        }

        public Guid? ContactId
        {
            get
            {
                if (_targetEntity["new_contact"] == null)
                    return null;
                return _targetEntity.GetAttributeValue<EntityReference>("new_contact").Id;
            }
        }

        public DateTime? ImportDate
        {
            get { return _targetEntity.GetAttributeValue<DateTime?>("new_projectimportdate"); }
        }

        public Guid Id { get { return _targetEntity.Id; } }

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

        public string Email
        {
            get { return _targetEntity.GetAttributeValue<string>("emailaddress1"); }
            set { _targetEntity["emailaddress1"] = value; }
        }

        public string ProjectName
        {
            get { return _targetEntity.GetAttributeValue<string>("new_projectname"); }
        }

        public Guid? AccountId
        {
            get
            {
                if (_targetEntity.Contains("new_account"))
                    return _targetEntity.GetAttributeValue<EntityReference>("new_account").Id;
                return null;
            }
        }

        public string AccountName
        {
            get { return _targetEntity.GetAttributeValue<string>("companyname"); }
        }

        public EntityReference ToRef()
        {
            return _targetEntity.ToEntityReference();
        }

        public Entity ToEntity()
        {
            return _targetEntity;
        }
    }
}
