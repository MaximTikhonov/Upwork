using MergeExcelDuplicates.Logic.Implementation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.ProxyObjects
{
    public class InitialRecordProxy
    {
        private Dictionary<global::MergeExcelDuplicates.Logic.Implementation.DataConnector.Columns, string> _rowData;

        public InitialRecordProxy(Dictionary<global::MergeExcelDuplicates.Logic.Implementation.DataConnector.Columns, string> rowData)
        {
            // TODO: Complete member initialization
            this._rowData = rowData;
        }
        internal AccountProxy ToAccount()
        {
            var accountResult = new AccountProxy()
            {
                Name = _rowData[DataConnector.Columns.ArchitectCompanyName],
                Address1 = _rowData[DataConnector.Columns.ArchitectCompanyAddress1],
                Address2 = _rowData[DataConnector.Columns.ArchitectCompanyAddress2],
                Address3 = _rowData[DataConnector.Columns.ArchitectCompanyAddress3],
                TownName = _rowData[DataConnector.Columns.ArchitectCompanyTown],
                CountyName = _rowData[DataConnector.Columns.ArchitectCounty],
                PostCode = _rowData[DataConnector.Columns.ArchitectPostCode],
                ImportDate = _rowData[DataConnector.Columns.DateOfInfoReceipt],
                ImportId = Guid.NewGuid()
            };
            return accountResult;
        }

        internal ProjectProxy ToProject()
        {
            var projectResult = new ProjectProxy()
            {
                Name = _rowData[DataConnector.Columns.ProjectName],
                ProjectSiteAddress1 = _rowData[DataConnector.Columns.ProjectSiteAddress1],
                ProjectSiteAddress2 = _rowData[DataConnector.Columns.ProjectSiteAddress2],
                ProjectSiteTown = _rowData[DataConnector.Columns.ProjectSiteTown],
                ProjectSiteCounty = _rowData[DataConnector.Columns.ProjectSiteCounty],
                ProjectSitePostcode = _rowData[DataConnector.Columns.ProjectSitePostcode],
                FullData = _rowData[DataConnector.Columns.FullData],
                ImportDate = _rowData[DataConnector.Columns.DateOfInfoReceipt],
                GovermentReference = _rowData[DataConnector.Columns.GovermentReference],
                ShortWorkDescription1 = _rowData[DataConnector.Columns.ShortWorkDescription1],
                ShortWorkDescription2 = _rowData[DataConnector.Columns.ShortWorkDescription2],
                ImportId = Guid.NewGuid()
            };
            return projectResult;
        }

        internal ContactProxy ToContact()
        {
            var contactResult = new ContactProxy()
            {
                FirstName = _rowData[DataConnector.Columns.ArchitectFirstName],
                SecondName = _rowData[DataConnector.Columns.ArchitectSecondName],
                ArchitectWorkPhone = _rowData[DataConnector.Columns.ArchitectWorkPhone],
                ArchitectEmail = _rowData[DataConnector.Columns.ArchitectEmail],
                ArchitectPhone = _rowData[DataConnector.Columns.ArchitectPhone],
                ArchitectRole = _rowData[DataConnector.Columns.ArchitectRole],
                ArchitectWebsite = _rowData[DataConnector.Columns.ArchitectWebsite],
                ArchitectFax = _rowData[DataConnector.Columns.ArchitectFax],
                ArchitectSex = _rowData[DataConnector.Columns.ArchitectSex],
                ImportId = Guid.NewGuid()
            };
            return contactResult;
        }
    }
}
