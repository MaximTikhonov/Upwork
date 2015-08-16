using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace MergingAccountsActivity.ProxyObjects
{
    public class LeadProxy : ICrmEntity
    {
        private readonly Entity _targetEntity;

        public LeadProxy(Entity targetEntity)
        {
            _targetEntity = targetEntity;
        }

        public Guid ContactId {
            get
            {
                if (!_targetEntity.Contains("new_contact") ||
                    _targetEntity.GetAttributeValue<EntityReference>("new_contact") == null)
                    return Guid.Empty;
                return _targetEntity.GetAttributeValue<EntityReference>("new_contact").Id;
            }
        }

        public string Topic {
            get
            {
                return _targetEntity.GetAttributeValue<string>("subject");
            }
            set { }
        }

        public Guid Id {
            get { return _targetEntity.Id; }
        }

        public string ProjectName {
            get { return _targetEntity.GetAttributeValue<string>("new_projectname"); }
            set { _targetEntity["new_projectname"] = value; }
        }

        public DateTime? ImportDate
        {
            get { return _targetEntity.GetAttributeValue<DateTime?>("new_projectimportdate"); }
            set { _targetEntity["new_projectimportdate"] = value; }
        }

        public Guid AccountId
        {
            get
            {
                if (!_targetEntity.Contains("new_account") ||
                    _targetEntity.GetAttributeValue<EntityReference>("new_account") == null)
                    return Guid.Empty;
                return _targetEntity.GetAttributeValue<EntityReference>("new_account").Id;
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
