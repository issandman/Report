using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class Record
    {
        private Workbook workBook;
        private Worksheet workSheet;
        //报表变量
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //报表路径
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        //连接数据库
        static string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;
        string today, yestoday, workface;

        public Record() { }
        public void Start(GetTime gt) {
            today = gt.getDateToday();
            OpenExcel();
            ToExcel();

        }

        private void ToExcel()
        {
            string sqlquary = string.Format(@"SELECT * FROM 现场显现记录 WHERE 日期='{0}'",today);
            DataTable dt = GetDataTable(sqlquary);
            DataRow[] drs = dt.Select();
            workSheet_excel.Cells["C2"].PutValue(drs[0]["现场记录"].ToString());
            workSheet_excel.Cells["C3"].PutValue(drs[0]["微震对应"].ToString());
            workSheet_excel.Cells["D2"].PutValue(drs[0]["具体对应位置"].ToString());
            workSheet_excel.Cells["J2"].PutValue(drs[0]["现场详细描述"].ToString());
            workSheet_excel.Cells["R2"].PutValue(drs[0]["分析说明"].ToString());

            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }
        private DataTable GetDataTable(string sqlquary)
        {
            sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
            DataTable datatable = new DataTable();
            sqlDataAdapter.Fill(datatable);
            return datatable;
        }
        private void OpenExcel()
        {
            string FilePath = @"..\..\modelFile\现场记录模板.xlsx";
            workBook = new Workbook(FilePath);
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            if (workBook_excel.Worksheets["Sheet7"] != null)
                workSheet_excel = workBook_excel.Worksheets["Sheet7"];
            else{
                workBook_excel.Worksheets.Add("Sheet7");
                workSheet_excel = workBook_excel.Worksheets["Sheet7"];
            }

            workSheet_excel.Copy(workSheet);
        }
    }
}
