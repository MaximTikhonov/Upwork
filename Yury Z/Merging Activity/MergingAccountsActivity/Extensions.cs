using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MergingAccountsActivity
{
    public static class Extensions
    {

        public static IEnumerable<T> CrmQuery<T>(this IOrganizationService service, QueryExpression query)
            where T : Entity
        {
            var PageInfoSpecified = query.PageInfo != null &&
                                    (query.PageInfo.PageNumber != 0 || query.PageInfo.Count != 0);
            EntityCollection items;
            do
            {
                items = service.RetrieveMultiple(query);
                foreach (var i in items.Entities)
                    yield return (T) i;

                if (!PageInfoSpecified && items.MoreRecords)
                {
                    if (query.PageInfo == null || (query.PageInfo.PageNumber == 0 && query.PageInfo.Count == 0))
                        query.PageInfo = new PagingInfo() {PageNumber = 1};
                    // Increment the page number to retrieve the next page.
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = items.PagingCookie;
                }
            } while (!PageInfoSpecified && items.MoreRecords);
        }
    }
}
