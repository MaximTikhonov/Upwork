using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingListFactory
{
    public class EventNotificator
    {
        public event EventHandler ProcessStarted = delegate { };
        public event EventHandler<ImportEventArgs> InitialDataRead = delegate { };
        public event EventHandler<ImportEventArgs> RecordProcessed = delegate { };

        internal void InitialDataCompleted(int totalRecords)
        {
            InitialDataRead.Invoke(null, new ImportEventArgs() { TotalRecords = totalRecords, CurrentRecordNum = 0 });
        }

        internal void StartProcess()
        {
            ProcessStarted(null, null);
        }

        internal void DataProcessed(int currentNum, int totalRecordsCount)
        {
            RecordProcessed(null, new ImportEventArgs() { TotalRecords = totalRecordsCount, CurrentRecordNum = currentNum });
        }
    }
    public class ImportEventArgs : EventArgs
    {
        public int TotalRecords { get; set; }
        public int CurrentRecordNum { get; set; }
    }
}
