using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MarketingListFactory.DataLogic.Comparison;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Crm
{
    public abstract class EntityCrmProxy: ICompareEntity
    {
        protected readonly Entity _targetEntity;

        protected EntityCrmProxy(Entity targetEntity)
        {
            _targetEntity = targetEntity;
        }

        public abstract bool IsEqual(ICompareEntity compareEntity);

        public Entity ToEntity()
        {
            return _targetEntity;
        }

        public EntityReference ToRef()
        {
            return _targetEntity.ToEntityReference();
        }

        public Guid Id { get { return _targetEntity.Id; } }
    }
}
