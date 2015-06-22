using MergeExcelDuplicates.ProxyObjects;
using MergeExcelDuplicates.ProxyObjects.InitialDataSource;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.Logic.Interfaces
{
    interface IDataConnector
    {
        List<InitialRecordProxy> LoadInitialData(System.IO.Stream sourceFileStream);

        void CreateFile(ICommonProxy[] projects, string directoryName, string[] columns, string fileName);
        //create deadlock

        InitialDataSheet LoadUnformattedData(System.IO.Stream sourceFile);

        void CreateFormattedFile(InitialDataSheet unformattedDataSheet);
    }
}
