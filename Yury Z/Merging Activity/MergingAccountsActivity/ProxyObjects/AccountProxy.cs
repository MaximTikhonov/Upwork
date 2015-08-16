using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;

namespace MergingAccountsActivity.ProxyObjects
{
    public class AccountProxy: ICrmEntity
    {
        private readonly Entity _targetEntity;

        public AccountProxy(Entity targetEntity)
        {
            _targetEntity = targetEntity;
        }

        public string PostCode
        {
            get { return _targetEntity.GetAttributeValue<string>("address1_postalcode").Replace(" ",""); }
        }

        public string ShortName {
            get { return _targetEntity.GetAttributeValue<string>("new_shortname"); }
        }

        public Guid Id {
            get { return _targetEntity.Id; }
        }

        public DateTime CreatedOn
        {
            get { return _targetEntity.GetAttributeValue<DateTime>("createdon"); }
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
