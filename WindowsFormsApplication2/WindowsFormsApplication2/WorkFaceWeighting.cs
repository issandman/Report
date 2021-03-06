﻿using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class WorkFaceWeighting
    {
        private static string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
        //模板excel文件
        private Workbook workBook;
        //模板工作sheet
        private Worksheet workSheet;
        //报表excel 
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //总报表excel 单例模式
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;

        public WorkFaceWeighting() { }
        public void Start(GetTime gt) {
            openexcel();
            string sql = string.Format(@"SELECT * FROM 工作面来压情况  WHERE 日期='{0}' AND 工作面='{1}' ", gt.getDateToday(), gt.getWorkFace());
            DataTable datatable = GetDataTable(sql);
            DataRow dr = datatable.Rows[0];
            if (dr != null) {
                Cell cellItem1 = workSheet_excel.Cells["C2"];
                Cell cellItem2 = workSheet_excel.Cells["D2"];
                Cell cellItem3 = workSheet_excel.Cells["E2"];
                Cell cellItem4 = workSheet_excel.Cells["F2"];
                Cell cellItem5 = workSheet_excel.Cells["G2"];
                Cell cellItem6 = workSheet_excel.Cells["I2"];
                Cell cellItem7 = workSheet_excel.Cells["C4"];
                Cell cellItem8 = workSheet_excel.Cells["E4"];
                Cell cellItem9 = workSheet_excel.Cells["H4"];
                Cell cellItem10 = workSheet_excel.Cells["C5"];
                Cell cellItem11 = workSheet_excel.Cells["H5"];


                cellItem1.PutValue(dr["已来压次数"].ToString());
                if (dr["上次位置"].ToString() != "0")
                    cellItem2.PutValue(dr["上次位置"].ToString());
                cellItem3.PutValue(dr["上次时间"].ToString());
                if (dr["步距"].ToString() != "0")
                    cellItem4.PutValue(dr["步距"].ToString());
                cellItem5.PutValue(dr["本次来压情况"].ToString());
                if (dr["持续距离"].ToString() != "0")
                    cellItem6.PutValue(dr["持续距离"].ToString());
                cellItem7.PutValue(dr["预计下次时间"].ToString());
                if (dr["预计下次位置"].ToString() != "0")
                    cellItem8.PutValue(dr["预计下次位置"].ToString());
                if (dr["预计下次步距"].ToString() != "0")
                    cellItem9.PutValue(dr["预计下次步距"].ToString());
                cellItem10.PutValue(dr["下一危险区域名称"].ToString());
                if (dr["距离危险区域"].ToString() != "0")
                    cellItem11.PutValue(dr["距离危险区域"].ToString());


            }

            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        private void openexcel()
        {
            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();
            string filepath = @"..\..\modelFile\工作面来压模板.xlsx";
            workBook = new Workbook(filepath);
            workSheet = workBook.Worksheets[0];

            if (workBook_excel.Worksheets["Sheet3"] != null)
            {
                workSheet_excel = workBook_excel.Worksheets["Sheet3"];
            }
            else
            {

                workBook_excel.Worksheets.Add("Sheet3");
                workSheet_excel = workBook_excel.Worksheets["Sheet3"];

            }

            workSheet_excel.Copy(workSheet);
        }

        private DataTable GetDataTable(string sqlquary)
        {
            DataTable datatable = new DataTable();
            sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
            sqlDataAdapter.Fill(datatable);

            return datatable;
        }
    }
}
