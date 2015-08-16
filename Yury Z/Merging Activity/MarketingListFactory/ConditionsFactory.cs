using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;

namespace MarketingListFactory
{
    public class ConditionsFactory
    {
        public static QueryExpression MarketingLists(string[] columns)
        {
            var result = new QueryExpression("list");
            result.ColumnSet = new ColumnSet(columns);
            result.Criteria.AddCondition("statecode", ConditionOperator.Equal,0);
            return result;
        }
    }
}
