using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.ProxyObjects.Crm;
using MarketingListFactory.ProxyObjects.Excel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using MarketingListFactory.ProxyObjects;

namespace MarketingListFactory.DataLogic
{
    public class CrmConnector
    {
        private readonly IOrganizationService _crmService;

        public CrmConnector(IOrganizationService crmService)
        {
            _crmService = crmService;
        }

        public List<AccountCrmProxy> GetAccounts(string[] columns, List<ConditionExpression> conditions)
        {
            var query=new QueryExpression("account");
            query.ColumnSet=new ColumnSet(columns);
            conditions.ForEach(p=>query.Criteria.AddCondition(p));
            return _crmService.CrmQuery<Entity>(query).Select(p => new AccountCrmProxy(p)).ToList();
        }

        public void MergeAccount(AccountCrmProxy target, AccountCrmProxy source)
        {
            var request=new MergeRequest
            {
                Target = target.ToRef(),
                SubordinateId = source.Id,
                PerformParentingChecks = false,
                UpdateContent = new Entity("account")
            };
            _crmService.Execute(request);
        }
       
        public Entity GetRecursiveEntity(EntityReference recursiveEntityRef, string[] columns)
        {
            return _crmService.Retrieve("new_recursivejob", recursiveEntityRef.Id, new ColumnSet(columns));
        }

        public void Update(Entity targetEntity)
        {
            _crmService.Update(targetEntity);
        }

        public List<ContactCrmProxy> GetContacts(string[] columns, List<ConditionExpression> conditions, LogicalOperator logicalOperator=LogicalOperator.And)
        {
            var query = new QueryExpression("contact");
            query.ColumnSet = new ColumnSet(columns);
            query.LinkEntities.Add(new LinkEntity("contact", "account", "parentcustomerid", "accountid", JoinOperator.Inner));
            query.LinkEntities[0].Columns.AddColumn("address1_postalcode");
            query.LinkEntities[0].EntityAlias = "parentcustomer";
            conditions.ForEach(p => query.Criteria.AddCondition(p));
            query.Criteria.FilterOperator = logicalOperator;
            return _crmService.CrmQuery<Entity>(query).Select(p => new ContactCrmProxy(p)).ToList();
        }

        //public void MergeContact(ContactExcelProxy target, ContactExcelProxy source)
        //{
        //    var updateEntity = new ContactExcelProxy(new Entity("contact"));
        //    if (target.FirstName.Length < source.FirstName.Length)
        //        updateEntity.FirstName = source.FirstName;
        //    if (target.Email == null && source.Email != null)
        //        updateEntity.Email = source.Email;
        //    if (target.Phone == null && source.Phone != null)
        //        updateEntity.Phone = source.Phone;
        //    if (target.Role == null && source.Role != null)
        //        updateEntity.Role = source.Role;

        //    var request = new MergeRequest
        //    {
        //        Target = target.ToEntity().ToEntityReference(),
        //        SubordinateId = source.Id,
        //        PerformParentingChecks = false,
        //        UpdateContent = updateEntity.ToEntity()
        //    };
        //    _crmService.Execute(request);
        //}

        public List<LeadCrmProxy> GetLeads(string[] columns, List<ConditionExpression> conditions)
        {
            var query = new QueryExpression("lead");
            query.ColumnSet = new ColumnSet(columns);
            conditions.ForEach(p => query.Criteria.AddCondition(p));
            return _crmService.CrmQuery<Entity>(query).Select(p => new LeadCrmProxy(p)).ToList();
        }

        //public void MergeLead(LeadProxy target, LeadProxy source)
        //{
        //    var updateEntity = new LeadProxy(new Entity("lead"));
        //    if ((target.ProjectName == null || target.ProjectName == target.Topic) && source.ProjectName != null)
        //        updateEntity.ProjectName = source.ProjectName;
        //    if (target.ImportDate == null && source.ImportDate != null)
        //        updateEntity.ImportDate = source.ImportDate;
            
        //    var request = new MergeRequest
        //    {
        //        Target = target.ToEntity().ToEntityReference(),
        //        SubordinateId = source.Id,
        //        PerformParentingChecks = false,
        //        UpdateContent = updateEntity.ToEntity()
        //    };
        //    _crmService.Execute(request);
        //}

        public Guid Create(Entity targetEntity)
        {
            return _crmService.Create(targetEntity);
        }
    }
}
