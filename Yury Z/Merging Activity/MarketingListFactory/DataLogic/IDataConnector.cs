using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.ProxyObjects;
using MarketingListFactory.ProxyObjects.Excel;

namespace MarketingListFactory.DataLogic
{
    interface IDataConnector
    {
        List<InitialRecordProxy> LoadInitialData();
        List<InitialRecordProxy> LoadInitialData(Dictionary<DataConnector.Columns, int> headers);

        void CreateFile(IEntityExcelProxy[] projects, string directoryName, string[] columns, string fileName);
        //create deadlock
        List<string> GetHeaderValues();
        InitialDataSheet LoadUnformattedData(System.IO.Stream sourceFile);

        void CreateFormattedFile(InitialDataSheet unformattedDataSheet);
        
    }
}
