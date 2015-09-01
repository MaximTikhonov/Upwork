using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.DataLogic;
using OfficeOpenXml;

namespace MarketingListFactory.ProxyObjects.Excel
{
    public class InitialDataSheet
    {
        private ExcelWorksheet workSheet;
        private static int startColumn = 15;
        private static int firstRoleColumn = 19;

        public InitialDataSheet(ExcelWorksheet workSheet)
        {
            // TODO: Complete member initialization
            this.workSheet = workSheet;            
        }

        internal int CountRows()
        {
            return workSheet.Dimension.End.Row;
        }

        internal void FormatRow(int rowIndex)
        {
            int tempInt;
            int companyNoIndex = -1;
            for (int i = startColumn; i <= workSheet.Dimension.End.Column; i++)
			{
                if (int.TryParse(workSheet.Cells[rowIndex, i].SafeStringValue(),out tempInt))
                {
                    companyNoIndex = i;
                    break;
                }
			}
            MergeUnecessaryCells(rowIndex,companyNoIndex);
        }

        private void MergeUnecessaryCells(int rowIndex, int companyNoIndex)
        {
            var difference=companyNoIndex-firstRoleColumn-1;
            if (difference <= 0)
                return;
            for (int i = firstRoleColumn; i <= workSheet.Dimension.End.Column; i++)
			{
                workSheet.Cells[rowIndex, i].Value = workSheet.Cells[rowIndex, i + difference].SafeStringValue();
			}            
        }

        internal ExcelWorksheet GetWorkSheet()
        {
            return workSheet;
        }
    }
}
