using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MergeExcelDuplicates.Logic.Implementation
{
    public class DataConnectorSimple: DataConnector
    {
        protected override Dictionary<Columns, int> ReadHeaders(ExcelWorksheet workSheet)
        {
            var result = new Dictionary<Columns, int>();
            var lastColumnIndex = workSheet.Dimension.End.Column;
            for (int rowIndex = 1; rowIndex <= 4; rowIndex++)
            {
                for (int columnIndex = 1; columnIndex <= lastColumnIndex; columnIndex++)
                {
                    switch (workSheet.Cells[rowIndex, columnIndex].SafeStringValue())
                    {
                        case "ID":
                            result[Columns.CustomId] = columnIndex;
                            break;
                        case "Project Number":
                            result[Columns.ProjectNo] = columnIndex;
                            break;
                        case "Column1"://Id of information supplier
                            result[Columns.IdOfInformationSupplier] = columnIndex;
                            break;
                        case "Full Data":
                            result[Columns.FullData] = columnIndex;
                            break;
                        case "Import Date":
                            result[Columns.DateOfInfoReceipt] = columnIndex;
                            break;
                        case "ApNo":
                            result[Columns.InternalId] = columnIndex;
                            break;
                        case "ExtRef":
                            result[Columns.GovermentReference] = columnIndex;
                            break;
                        case "AppDate":
                            result[Columns.AppDate] = columnIndex;
                            break;
                        case "Client Title":
                            result[Columns.ClientTitle] = columnIndex;
                            break;
                        case "Client Name":
                            result[Columns.ClientName] = columnIndex;
                            break;
                        case "Client Surname":
                            result[Columns.ClientSurname] = columnIndex;
                            break;
                        case "Client Company Name":
                            result[Columns.ClientCompanyName] = columnIndex;
                            break;
                        case "Client Address 1":
                            result[Columns.ClientAddress1] = columnIndex;
                            break;
                        case "Client Address 2":
                            result[Columns.ClientAddress2] = columnIndex;
                            break;
                        case "Client Address 3":
                            result[Columns.ClientAddress3] = columnIndex;
                            break;
                        case "Client Town":
                            result[Columns.ClientTown] = columnIndex;
                            break;
                        case "Client County":
                            result[Columns.ClientCounty] = columnIndex;
                            break;
                        case "Client Postcode":
                            result[Columns.ClientPostcode] = columnIndex;
                            break;
                        case "Description 1":
                            result[Columns.ShortWorkDescription1] = columnIndex;
                            break;
                        case "Description 2":
                            result[Columns.ShortWorkDescription2] = columnIndex;
                            break;
                        case "Title2":
                            result[Columns.ArchitectSex] = columnIndex;
                            break;
                        case "First Name2":
                            result[Columns.ArchitectFirstName] = columnIndex;
                            break;
                        case "Second Name3":
                            result[Columns.ArchitectLastName] = columnIndex;
                            break;
                        case "company_name":
                            result[Columns.ArchitectCompanyName] = columnIndex;
                            break;
                        case "Address 15":
                            result[Columns.ArchitectCompanyAddress1] = columnIndex;
                            break;
                        case "Address 26":
                            result[Columns.ArchitectCompanyAddress2] = columnIndex;
                            break;
                        case "Address 37":
                            result[Columns.ArchitectCompanyAddress3] = columnIndex;
                            break;
                        case "Town8":
                            result[Columns.ArchitectCompanyTown] = columnIndex;
                            break;
                        case "County9":
                            result[Columns.ArchitectCounty] = columnIndex;
                            break;
                        case "Post Code10":
                            result[Columns.ArchitectPostCode] = columnIndex;
                            break;
                        case "Work Phone Cleaned1":
                            result[Columns.ArchitectWorkPhone] = columnIndex;
                            break;
                        case "Lead Name12":
                            result[Columns.ProjectName] = columnIndex;
                            break;
                        case "Site Address 113":
                            result[Columns.ProjectSiteAddress1] = columnIndex;
                            break;
                        case "Site Address 214":
                            result[Columns.ProjectSiteAddress2] = columnIndex;
                            break;
                        case "Site Town15":
                            result[Columns.ProjectSiteTown] = columnIndex;
                            break;
                        case "Site County16":
                            result[Columns.ProjectSiteCounty] = columnIndex;
                            break;
                        case "Site Postcode17":
                            result[Columns.ProjectSitePostcode] = columnIndex;
                            break;
                        case "DevType20":
                            result[Columns.ProjectDevelopmentType] = columnIndex;
                            break;
                        case "Email 1":
                            result[Columns.ArchitectEmail] = columnIndex;
                            break;
                        case "Client phone Cleaned":
                            result[Columns.ClientCleanedPhone] = columnIndex;
                            break;
                        //case "Work Phone Cleaned":
                        //    result[Columns.ArchitectPhone] = columnIndex;
                        //    break;
                        case "Role1":
                            result[Columns.ArchitectRole] = columnIndex;
                            break;
                        case "Site1":
                            result[Columns.ArchitectWebsite] = columnIndex;
                            break;
                        case "Fax1":
                            result[Columns.ArchitectFax] = columnIndex;
                            break;
                    }
                }
                if (result.Any())
                    break;
            }
            return result;
        }
    }
}
