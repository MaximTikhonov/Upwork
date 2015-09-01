using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingListFactory.ProxyObjects.Common
{
    public interface IAccountProxy
    {
        string Address1 { get; set; }
        string Address2 { get; set; }
        string Address3 { get; set; }
        string CountyName { get; set; }
        string Email { get; set; }
        string ExtraPhones { get; set; }
        string ImportDate { get; set; }
        Guid ImportId { get; set; }
        string MobilePhone { get; set; }
        string Name { get; set; }
        string PostCode { get; set; }
        string ShortName { get; set; }
        string TownName { get; set; }
        string WebSite { get; set; }
        string WorkPhone { get; set; }

    }
}
