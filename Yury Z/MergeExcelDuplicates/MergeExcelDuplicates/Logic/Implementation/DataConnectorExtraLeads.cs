using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MergeExcelDuplicates.Logic.Implementation
{
    public class DataConnectorExtraLeads: DataConnector
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
                        case "ContactTitle":
                            result[Columns.ArchitectSex] = columnIndex;
                            break;
                        case "ContactFirst_name":
                            result[Columns.ArchitectFirstName] = columnIndex;
                            break;
                        case "ContactLast_name":
                            result[Columns.ArchitectLastName] = columnIndex;
                            break;
                        case "company_name":
                            result[Columns.ArchitectCompanyName] = columnIndex;
                            break;
                        case "company_address1":
                            result[Columns.ArchitectCompanyAddress1] = columnIndex;
                            break;
                        case "company_address2":
                            result[Columns.ArchitectCompanyAddress2] = columnIndex;
                            break;
                        //case "company_address3":
                        //    result[Columns.ArchitectCompanyAddress3] = columnIndex;
                        //    break;
                        case "company_town":
                            result[Columns.ArchitectCompanyTown] = columnIndex;
                            result[Columns.ArchitectCompanyAddress3] = columnIndex;
                            break;
                        case "company_county":
                            result[Columns.ArchitectCounty] = columnIndex;
                            break;
                        case "postcode":
                            result[Columns.ArchitectPostCode] = columnIndex;
                            break;
                        case "company_phone":
                            result[Columns.ArchitectWorkPhone] = columnIndex;
                            result[Columns.ArchitectPhone] = columnIndex;
                            break;
                        case "company_email":
                            result[Columns.ArchitectCompanyEmail] = columnIndex;
                            break;
                        case "ContactEmail_address":
                            result[Columns.ArchitectEmail] = columnIndex;
                            break;
                        case "Client phone Cleaned":
                            result[Columns.ClientCleanedPhone] = columnIndex;
                            break;
                        case "Work Phone Cleaned":
                            result[Columns.ArchitectPhone] = columnIndex;
                            break;
                        case "ContactRole_name":
                            result[Columns.ArchitectRole] = columnIndex;
                            break;
                        case "Site":
                            result[Columns.ArchitectWebsite] = columnIndex;
                            break;
                        case "Fax":
                            result[Columns.ArchitectFax] = columnIndex;
                            break;
                        //Project
                        case "Project_no":
                            result[Columns.ProjectNo] = columnIndex;
                            break;
                        case "Project_title":
                            result[Columns.ProjectName] = columnIndex;
                            result[Columns.FullData] = columnIndex;
                            break;
                        case "ProjectShort_site_address":
                            result[Columns.ProjectSiteAddress1] = columnIndex;
                            break;
                        case "ProjectSite_address1":
                            result[Columns.ProjectSiteAddress2] = columnIndex;
                            break;
                        case "Site Town":
                            result[Columns.ProjectSiteTown] = columnIndex;
                            break;
                        case "Site County":
                            result[Columns.ProjectSiteCounty] = columnIndex;
                            break;
                        case "Project_value":
                            result[Columns.ProjectValue] = columnIndex;
                            break;
                        case "ProjectPlanning_ref":
                            result[Columns.ProjectPlanningRef] = columnIndex;
                            break;
                        case "ProjectDevelopment type":
                            result[Columns.ProjectDevelopmentType] = columnIndex;
                            break;
                        case "ProjectStage":
                            result[Columns.ProjectStage] = columnIndex;
                            break;
                        case "ProjectStatus":
                            result[Columns.ProjectStatus] = columnIndex;
                            break;
                        case "ProjectScheme details":
                            result[Columns.ProjectSchemeDetails] = columnIndex;
                            break;
                        case "ProjectProgramme timing":
                            result[Columns.ProjectProgrammeTiming] = columnIndex;
                            break;
                        case "ProjectContract type":
                            result[Columns.ProjectContractType] = columnIndex;
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
