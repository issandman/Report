using Aspose.Cells;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class MergeTest
    {
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();

        public void MergeAllSheet(GetTime gt ) {

            Spire.Xls.Workbook wb = new Spire.Xls.Workbook();
            wb.LoadFromFile(excelFilePath);
            Spire.Xls.Workbook destWorkbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet destSheet = destWorkbook.Worksheets[0];
            for (int i = 1; i <= wb.Worksheets.Count + 6; i++) {
                if (wb.Worksheets["Sheet" + i.ToString()] != null)
                    wb.Worksheets["Sheet" + i.ToString()].AllocatedRange.Copy(destSheet.Range[destSheet.LastRow + 3, 1]);
            }
            destWorkbook.SaveToFile(@"..\..\modelFile\" + string.Format("Excel_{0}.xls", gt.getDateToday()));
        }
    }
}
