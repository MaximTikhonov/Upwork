using MergeExcelDuplicates.Logic.Interfaces;
using MergeExcelDuplicates.ProxyObjects;
using MergeExcelDuplicates.ProxyObjects.InitialDataSource;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.Logic.Implementation
{
    public class DataConnector : IDataConnector
    {
        private int headersRow;
        public enum Columns
        {
            CustomId,
            IdOfInformationSupplier,
            FullData,//project
            DateOfInfoReceipt,//project company
            InternalId,
            GovermentReference,//project
            AppDate,
            ClientTitle,
            ClientName,
            ClientSurname,
            ClientCompanyName,
            ClientAddress1,
            ClientAddress2,
            ClientAddress3,
            ClientTown,
            ClientCounty,
            ClientPostcode,
            ShortWorkDescription1,//project
            ShortWorkDescription2,//project
            ArchitectSex,//contact
            ArchitectFirstName,//contact
            ArchitectSecondName,//contact
            ArchitectCompanyName,//company
            ArchitectCompanyAddress1,//company
            ArchitectCompanyAddress2,//company
            ArchitectCompanyAddress3,//company
            ArchitectCompanyTown,//company
            ArchitectCounty,//company
            ArchitectPostCode,//company
            ArchitectWorkPhone,//contact
            ProjectName,//project
            ProjectSiteAddress1,//project
            ProjectSiteAddress2,//project
            ProjectSiteTown,//project
            ProjectSiteCounty,//project
            ProjectSitePostcode,//project
            ArchitectEmail,//contact
            ClientCleanedPhone,
            ArchitectPhone,//contact
            ArchitectRole,//contact
            ArchitectWebsite,//contact
            ArchitectFax//contact
        }
        public List<InitialRecordProxy> LoadInitialData(Stream sourceFileStream)
        {
            var sourceExcelPackage = new ExcelPackage(sourceFileStream);
            var dataSheetName = MergeExcelDuplicates.Properties.Settings.Default.ExcelWorksheetName;
            var workSheet = sourceExcelPackage.Workbook.Worksheets[dataSheetName];
            var headers = ReadHeaders(workSheet);
            if (!headers.Any())
                throw new Exception(MergeExcelDuplicates.Properties.Resources.ExceptionEmptyHeaders);
            var endRow = workSheet.Dimension.End.Row;
            var result = new List<InitialRecordProxy>();
            for (int rowIndex = 2; rowIndex <= endRow; rowIndex++)
            {
                var initialRecord = GetRowRecord(rowIndex, headers, workSheet);
                result.Add(initialRecord);
            }
            return result;
        }

        private InitialRecordProxy GetRowRecord(int rowIndex, Dictionary<Columns, int> headers, ExcelWorksheet workSheet)
        {
            var rowData = new Dictionary<Columns, string>();
            var columnsCount = workSheet.Dimension.End.Column;
            for (int columnIndex = 1; columnIndex <= columnsCount; columnIndex++)
            {
                if (!headers.Any(p => p.Value == columnIndex))
                    continue;
                var columnType = headers.FirstOrDefault(p => p.Value == columnIndex).Key;
                rowData.Add(columnType, workSheet.Cells[rowIndex, columnIndex].SafeStringValue());
            }
            var result = new InitialRecordProxy(rowData);
            return result;
        }

        private Dictionary<Columns, int> ReadHeaders(ExcelWorksheet workSheet)
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
                        case "Title":
                            result[Columns.ArchitectSex] = columnIndex;
                            break;
                        case "First Name":
                            result[Columns.ArchitectFirstName] = columnIndex;
                            break;
                        case "Second Name":
                            result[Columns.ArchitectSecondName] = columnIndex;
                            break;
                        case "company_name":
                            result[Columns.ArchitectCompanyName] = columnIndex;
                            break;
                        case "Address 1":
                            result[Columns.ArchitectCompanyAddress1] = columnIndex;
                            break;
                        case "Address 2":
                            result[Columns.ArchitectCompanyAddress2] = columnIndex;
                            break;
                        case "Address 3":
                            result[Columns.ArchitectCompanyAddress3] = columnIndex;
                            break;
                        case "Town":
                            result[Columns.ArchitectCompanyTown] = columnIndex;
                            break;
                        case "County":
                            result[Columns.ArchitectCounty] = columnIndex;
                            break;
                        case "Post Code":
                            result[Columns.ArchitectPostCode] = columnIndex;
                            break;
                        case "Work Phone":
                            result[Columns.ArchitectWorkPhone] = columnIndex;
                            break;
                        case "Lead Name":
                            result[Columns.ProjectName] = columnIndex;
                            break;
                        case "Site Address 1":
                            result[Columns.ProjectSiteAddress1] = columnIndex;
                            break;
                        case "Site Address 2":
                            result[Columns.ProjectSiteAddress2] = columnIndex;
                            break;
                        case "Site Town":
                            result[Columns.ProjectSiteTown] = columnIndex;
                            break;
                        case "Site County":
                            result[Columns.ProjectSiteCounty] = columnIndex;
                            break;
                        case "Site Postcode":
                            result[Columns.ProjectSitePostcode] = columnIndex;
                            break;
                        case "email":
                            result[Columns.ArchitectEmail] = columnIndex;
                            break;
                        case "Client phone Cleaned":
                            result[Columns.ClientCleanedPhone] = columnIndex;
                            break;
                        case "Work Phone Cleaned":
                            result[Columns.ArchitectPhone] = columnIndex;
                            break;
                        case "Role":
                            result[Columns.ArchitectRole] = columnIndex;
                            break;
                        case "Site":
                            result[Columns.ArchitectWebsite] = columnIndex;
                            break;
                        case "Fax":
                            result[Columns.ArchitectFax] = columnIndex;
                            break;
                    }
                }
                if (result.Any())
                    break;
            }
            return result;
        }


        public void CreateFile(ICommonProxy[] source, string directoryName, string[] columns, string fileName)
        {
            if (File.Exists(directoryName + "Result//" + fileName))
            {
                File.Delete(directoryName + "Result//" + fileName);
            }
            var fileInfo = new FileInfo(directoryName + "Result//" + fileName);
            using (var package = new ExcelPackage(fileInfo))
            {
                var targetWorksheet = package.Workbook.Worksheets.Add("ImportResult");
                for (int i = 0; i < columns.Count(); i++)
                {
                    targetWorksheet.Cells[1, i + 1].Value = columns[i];
                }
                for (int sourceIndex = 0; sourceIndex < source.Count(); sourceIndex++)
                {
                    var sourceItem = source[sourceIndex];
                    var sourceItemData = sourceItem.GetData();
                    for (int fieldIndex = 0; fieldIndex < columns.Count(); fieldIndex++)
                    {
                        var fieldName = columns[fieldIndex];
                        targetWorksheet.Cells[sourceIndex + 2, fieldIndex + 1].Value = sourceItemData.FirstOrDefault(p => p.Key == fieldName).Value;
                    }

                }
                package.Save();
            }
        }


        public InitialDataSheet LoadUnformattedData(Stream sourceFile)
        {
            var sourceExcelPackage = new ExcelPackage(sourceFile);
            var dataSheetName = MergeExcelDuplicates.Properties.Settings.Default.ExcelWorksheetName;
            var workSheet = sourceExcelPackage.Workbook.Worksheets[dataSheetName];
            var result = new InitialDataSheet(workSheet);
            return result;
        }


        public void CreateFormattedFile(InitialDataSheet formattedDataSheet)
        {
            if (File.Exists("Result//FormattedData.xlsx"))
            {
                File.Delete("Result//FormattedData.xlsx");
            }
            var fileInfo = new FileInfo("Result//FormattedData.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                package.Workbook.Worksheets.Add("Sheet1", formattedDataSheet.GetWorkSheet());                
                package.Save();
            }
        }
    }
}
