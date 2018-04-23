using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication2
{
    class Comprehensive
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
        //查询唯一的channel 和 查询value、location
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;
        string today, yestoday, workface;

        public Comprehensive() {
                
        }
        public void Start(GetTime gt)
        {
            today = gt.getDateToday();
            OpenExcel();
            ToExcel();

        }

        private void ToExcel()
        {
            string sqlquary = string.Format(@"SELECT * FROM 综合分析 WHERE 日期='{0}'", today);
            DataTable dt = GetDataTable(sqlquary);
            DataRow[] drs = dt.Select();
            workSheet_excel.Cells["B1"].PutValue(drs[0]["综合分析"].ToString());
            string way = drs[0]["采取措施"].ToString();
            Debug.WriteLine(way);
            string[] ways = Regex.Split(way, ", ", RegexOptions.IgnoreCase);
            string ss = null;
            for (int i = 0; i < ways.Length; i++) {
                ss += string.Format("{0}.{1}\n",i + 1, ways[i]) ;

            }
            workSheet_excel.Cells["E1"].PutValue(ss);
            workSheet_excel.Cells["L2"].PutValue(drs[0]["轨顺钻孔施工地点"].ToString());
            workSheet_excel.Cells["L3"].PutValue(drs[0]["胶运钻孔施工地点"].ToString());
            workSheet_excel.Cells["P2"].PutValue(drs[0]["轨顺孔深"].ToString());
            workSheet_excel.Cells["P3"].PutValue(drs[0]["胶运孔深"].ToString());
            workSheet_excel.Cells["Q2"].PutValue(drs[0]["轨顺孔间距"].ToString());
            workSheet_excel.Cells["Q3"].PutValue(drs[0]["胶运孔间距"].ToString());
            workSheet_excel.Cells["R2"].PutValue(drs[0]["轨顺设计个数"].ToString());
            workSheet_excel.Cells["S2"].PutValue(drs[0]["轨顺当日施工个数"].ToString());
            workSheet_excel.Cells["T2"].PutValue(drs[0]["轨顺剩余个数"].ToString());
            workSheet_excel.Cells["R3"].PutValue(drs[0]["胶运设计个数"].ToString());
            workSheet_excel.Cells["S3"].PutValue(drs[0]["胶运当日施工个数"].ToString());
            workSheet_excel.Cells["T3"].PutValue(drs[0]["胶运剩余个数"].ToString());
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
            string FilePath = @"..\..\modelFile\综合分析模板.xlsx";
            workBook = new Workbook(FilePath);
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            if (workBook_excel.Worksheets["Sheet8"] != null)
                workSheet_excel = workBook_excel.Worksheets["Sheet8"];
            else
            {
                workBook_excel.Worksheets.Add("Sheet8");
                workSheet_excel = workBook_excel.Worksheets["Sheet8"];
            }

            workSheet_excel.Copy(workSheet);
        }
    }
}
