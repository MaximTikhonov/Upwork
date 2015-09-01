using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.ProxyObjects.Crm;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Excel
{
    public interface IEntityExcelProxy
    {
        IEnumerable<KeyValuePair<string, string>> GetData();

        bool EqualsToAccount(AccountExcelProxy targetAccountExcel);

        bool EqualsToContact(ContactExcelProxy targetContactExcel);

        bool EqualsToProject(ProjectExcelProxy targetProjectExcel);

    }
}
