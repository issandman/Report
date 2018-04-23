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
            wb.Worksheets["Sheet1"].AllocatedRange.Copy(destSheet.Range[1, 1]);
            destSheet.Range[1, 1].Text = string.Format("营盘壕煤矿{0}工作面矿压综合分析日报表",gt.getWorkFace());
            int lastrow = destSheet.Range.RowCount;
            if (wb.Worksheets["Sheet2"] != null)
            {
                wb.Worksheets["Sheet2"].AllocatedRange.Copy(destSheet.Range[destSheet.Range.RowCount + 1, 1]);
                if (wb.Worksheets["Sheet3"] != null)
                {
                    wb.Worksheets["Sheet3"].AllocatedRange.Copy(destSheet.Range[lastrow + 1, 11]);
                }
                for (int i = 4; i <= wb.Worksheets.Count + 6; i++)
                {
                    if (wb.Worksheets["Sheet" + i.ToString()] != null)
                        wb.Worksheets["Sheet" + i.ToString()].AllocatedRange.Copy(destSheet.Range[destSheet.Range.RowCount + 1, 1]);
                }
            }
            destWorkbook.SaveToFile(@"..\..\modelFile\" + string.Format("Excel_{0}.xls", gt.getDateToday()));
        }
    }
}
