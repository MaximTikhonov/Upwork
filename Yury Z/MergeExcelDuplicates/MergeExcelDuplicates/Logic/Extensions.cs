using MergeExcelDuplicates.ProxyObjects;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.Logic
{
    public static class Extensions
    {
        public static string SafeStringValue(this ExcelRange targetCell)
        {
            if (targetCell == null || targetCell.Value == null)
                return "";
            return targetCell.Value.ToString();
        }
        public static AccountProxy GetEqualAccount(this List<ICommonProxy> sourceList, AccountProxy targetAccount)
        {
            var result = (from resultAccount in sourceList
                          where resultAccount.EqualsToAccount(targetAccount)
                          select resultAccount).FirstOrDefault();
            if (result == null)
                return null;
            return (AccountProxy)result;
        }
        public static ContactProxy GetEqualContact(this List<ICommonProxy> sourceList, ContactProxy targetContact)
        {
            var result = (from resultContact in sourceList
                          where resultContact.EqualsToContact(targetContact)
                          select resultContact).FirstOrDefault();
            if (result == null)
                return null;
            return (ContactProxy)result;
        }
        public static ProjectProxy GetEqualProject(this List<ICommonProxy> sourceList, ProjectProxy targetProject)
        {
            var result = (from resultProject in sourceList
                          where resultProject.EqualsToProject(targetProject)
                          select resultProject).FirstOrDefault();
            if (result == null)
                return null;
            return (ProjectProxy)result;
        }
    }
}
