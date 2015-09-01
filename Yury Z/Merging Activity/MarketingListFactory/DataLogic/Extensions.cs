using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.ProxyObjects;
using MarketingListFactory.ProxyObjects.Excel;
using OfficeOpenXml;

namespace MarketingListFactory.DataLogic
{
    public static class Extensions
    {
        public static string SafeStringValue(this ExcelRange targetCell)
        {
            if (targetCell == null || targetCell.Value == null)
                return "";
            return targetCell.Value.ToString();
        }
        public static AccountExcelProxy GetEqualAccount(this List<AccountExcelProxy> sourceList, AccountExcelProxy targetAccountExcel)
        {
            var result = (from resultAccount in sourceList
                          where resultAccount.EqualsToAccount(targetAccountExcel)
                          select resultAccount).FirstOrDefault();
            if (result == null)
                return null;
            return result;
        }
        public static ContactExcelProxy GetEqualContact(this List<ContactExcelProxy> sourceList, ContactExcelProxy targetContactExcel)
        {
            var result = (from resultContact in sourceList
                          where resultContact.EqualsToContact(targetContactExcel)
                          select resultContact).FirstOrDefault();
            if (result == null)
                return null;
            return result;
        }
        public static ProjectExcelProxy GetEqualProject(this List<ProjectExcelProxy> sourceList, ProjectExcelProxy targetProjectExcel)
        {
            var result = (from resultProject in sourceList
                          where resultProject.EqualsToProject(targetProjectExcel)
                          select resultProject).FirstOrDefault();
            if (result == null)
                return null;
            return result;
        }
    }
}
