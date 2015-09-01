using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingListFactory.DataLogic.Comparison
{
    public interface ICompareEntity
    {
        bool IsEqual(ICompareEntity compareEntity);
    }
}
