using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Aspose.Cells;
/*
    1443955916@qq.com
*/
namespace WindowsFormsApplication2
{
    //单例模式
    class CreateExcelTest
    {
        Workbook workBook_excel;
        Worksheet workSheet_excel;
        private static CreateExcelTest createExcel;
        //string excelFilePath = @"C:\Users\14439\Desktop\yingpanhao\报表\"
        //                        + string.Format("Excel_{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd"));//getTime();
        string today = DateTime.Now.ToString("yyyy-MM-dd");
        //string excelFilePath= @"C:\Users\14439\Desktop\yingpanhao\报表\";
        string filepath = @"..\..\modelFile\今日分报表.xlsx";
        //string filepath = @"C:\Users\han\Desktop\报表\ExcelToday.xlsx";

        private CreateExcelTest()
        {
            //导入破解证书
            try
            {
                Excel.License el = new Excel.License();
                el.SetLicense("Aid/License.lic");
            }
            catch (Exception)
            {
                //...
            }

            
            workBook_excel = File.Exists(filepath) ? new Workbook(filepath) : new Workbook();
            workSheet_excel = workBook_excel.Worksheets["Sheet1"];
            try
            {
                Workbook wb = new Workbook(@"..\..\modelFile\表头模板.xlsx");//营盘壕煤矿2101工作面矿压综合分析日报表
                Worksheet ws = wb.Worksheets[0];
                workSheet_excel.Copy(ws);
                workBook_excel.Save(filepath, SaveFormat.Xlsx);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("文件已打开，请关闭文件后再试");
            }
        }

        public static CreateExcelTest GetCreateExcelTest()
        {
            if (createExcel == null)
                return createExcel = new CreateExcelTest();
            return createExcel;
        }


        public Workbook GetWorkBookExcel() {

            return workBook_excel;
        }
        public string ExcelFilePath()
        {
            return filepath;
        }
    }
}
