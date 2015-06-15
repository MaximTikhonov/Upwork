using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.ProxyObjects
{
    public interface ICommonProxy
    {
        IEnumerable<KeyValuePair<string, string>> GetData();

        bool EqualsToAccount(AccountProxy targetAccount);

        bool EqualsToContact(ContactProxy targetContact);

        bool EqualsToProject(ProjectProxy targetProject);

    }
}
