using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace MergingAccountsActivity
{
    public static class ExtensionsFetch
    {

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

                returnCollection = ((RetrieveMultipleResponse) _service.Execute(fetchRequest1)).EntityCollection;

                foreach (var c in returnCollection.Entities)
                {
                    yield return (T) c;
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
           
            //EntityCollection items;
            //do
            //{
            //    items = service.RetrieveMultiple(query);
            //    foreach (var i in items.Entities)
            //        yield return (T) i;

            //    if (items.MoreRecords)
            //    {
            //        if (query.PageInfo == null || (query.PageInfo.PageNumber == 0 && query.PageInfo.Count == 0))
            //            query.PageInfo = new PagingInfo() {PageNumber = 1};
            //        // Increment the page number to retrieve the next page.
            //        query.PageInfo.PageNumber++;
            //        query.PageInfo.PagingCookie = items.PagingCookie;
            //    }
            //} while (!PageInfoSpecified && items.MoreRecords);
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
