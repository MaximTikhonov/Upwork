using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;

namespace MarketingListFactory
{
    public static class Extensions
    {
        public static string SafeStringValue(this ExcelRange targetCell)
        {
            if (targetCell == null || targetCell.Value == null)
                return "";
            return targetCell.Value.ToString();
        }
        public static T GetEqualEntity<T>(this List<T> sourceList, T targetEntity) where T: ICompareEntity  
        {
            var result = (from resultEntity in sourceList
                          where resultEntity.IsEqual(targetEntity)
                          select resultEntity).FirstOrDefault();
            return (T)result;
        }
        public static bool PhoneCompare(string existingValue, string newValue)
        {
            if(string.IsNullOrEmpty(existingValue)&&string.IsNullOrEmpty(newValue))
                return true;
            if(string.IsNullOrEmpty(existingValue)||string.IsNullOrEmpty(newValue))
                return false;
            var exisitingPhone = phoneCompare.Match(existingValue).Value;
            var newPhone = phoneCompare.Match(newValue).Value;
            if (newPhone.Length < 6)
                return false;
            return exisitingPhone.Contains(newPhone);
        }
        private static Regex phoneCompare = new Regex(@"\d+");
        //public static AccountProxy GetEqualAccount(this List<ICommonProxy> sourceList, AccountProxy targetAccount)
        //{
        //    var result = (from resultAccount in sourceList
        //                  where resultAccount.EqualsToAccount(targetAccount)
        //                  select resultAccount).FirstOrDefault();
        //    if (result == null)
        //        return null;
        //    return (AccountProxy)result;
        //}
        //public static ContactProxy GetEqualContact(this List<ICommonProxy> sourceList, ContactProxy targetContact)
        //{
        //    var result = (from resultContact in sourceList
        //                  where resultContact.EqualsToContact(targetContact)
        //                  select resultContact).FirstOrDefault();
        //    if (result == null)
        //        return null;
        //    return (ContactProxy)result;
        //}
        //public static ProjectProxy GetEqualProject(this List<ICommonProxy> sourceList, ProjectProxy targetProject)
        //{
        //    var result = (from resultProject in sourceList
        //                  where resultProject.EqualsToProject(targetProject)
        //                  select resultProject).FirstOrDefault();
        //    if (result == null)
        //        return null;
        //    return (ProjectProxy)result;
        //}
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
        public static IEnumerable<T> CrmQuery<T>(this IOrganizationService _service, string fetchXml)
           where T : Entity
        {
            int fetchCount = 0;
            // Initialize the page number.
            int pageNumber = 0;
            // Initialize the number of records.
            int recordCount = 0;
            // Specify the current paging cookie. For retrieving the first page, 
            // pagingCookie should be null.
            string pagingCookie = null;
            EntityCollection returnCollection;
            do
            {
                string xml = CreateXml(fetchXml, pagingCookie, pageNumber, fetchCount);
                var fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                returnCollection = ((RetrieveMultipleResponse)_service.Execute(fetchRequest1)).EntityCollection;

                foreach (var c in returnCollection.Entities)
                {
                    yield return (T)c;
                }
                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {


                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            } while (returnCollection.MoreRecords);
        }
        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }
    }
}
