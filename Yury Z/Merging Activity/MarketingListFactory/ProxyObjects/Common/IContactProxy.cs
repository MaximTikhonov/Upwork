using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Resources;

namespace MarketingListFactory.ProxyObjects.Common
{
    public interface IContactProxy
    {
        string LastName { get; set; }
        string FirstName { get; set; }
    }
}
