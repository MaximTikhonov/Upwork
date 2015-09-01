using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarketingListFactory.DataLogic.Comparison
{
    public class FuzzyStringComparer
    {
        public static bool IsStringsFuzzyEquals(string first, string second, int equalityPercent)
        {
            if (string.IsNullOrWhiteSpace(second))
                return false;
            var computeResult = LevenshteinDistance.Compute(first, second);
            return 1 - (decimal)computeResult / second.Length > (decimal)equalityPercent / 100;
        }
    }
}
