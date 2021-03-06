﻿using MergeExcelDuplicates.Logic.Interfaces;
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
    public abstract class DataConnector : IDataConnector
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
            ArchitectLastName,//contact
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
            ,
            ArchitectCompanyEmail,
            ProjectNo,
            ProjectValue,
            ProjectPlanningRef,
            ProjectDevelopmentType,
            ProjectStage,
            ProjectStatus,
            ProjectSchemeDetails,
            ProjectProgrammeTiming,
            ProjectContractType
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
        abstract  protected Dictionary<DataConnector.Columns, int> ReadHeaders(ExcelWorksheet workSheet);
        protected InitialRecordProxy GetRowRecord(int rowIndex, Dictionary<Columns, int> headers, ExcelWorksheet workSheet)
        {
            var rowData = new Dictionary<Columns, string>();
            var columnsCount = workSheet.Dimension.End.Column;
            for (int columnIndex = 1; columnIndex <= columnsCount; columnIndex++)
            {
                if (!headers.Any(p => p.Value == columnIndex))
                    continue;
                var columnType = headers.Where(p => p.Value == columnIndex);
                foreach (var iHeader in columnType)
                {
                    if (!rowData.ContainsKey(iHeader.Key))
                    rowData.Add(iHeader.Key, workSheet.Cells[rowIndex, columnIndex].SafeStringValue());    
                }
                
            }
            var result = new InitialRecordProxy(rowData);
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
