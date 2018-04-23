using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class StandTest
    {
        private static string constr = "server=.;database=ylkdb;uid=sa;pwd=sakjdx";
        static string conn = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
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
        SqlConnection sqlConn = new SqlConnection(conn);
        SqlDataAdapter sqlDataAdapter;
        private string[] locations = { "上部", "中部", "下部" };
        private string[] pic_path = { @"..\..\modelFile\chart3.png", @"..\..\modelFile\chart4.png", @"..\..\modelFile\chart5.png" };
        string today, yestoday;
        double sum = 0;
        double ava = 0;
        double[] average = new double[3];
        double[] max_three = new double[3];
        string behaviors;
        public StandTest() {
            
        }
        public void Start(GetTime gt, double ava, string behaviors) {

            this.behaviors = behaviors;
            today = gt.getDateToday();
            yestoday = gt.getDateYestoday();
            this.ava = ava;
            OpenExcel();
            Calculate();
            ToExcel();
            AddPicures();
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
            sqlConn.Close();
        }

        private void AddPicures()
        {
            workSheet_excel.Pictures.Add(6, 1, 20, 5, pic_path[0]);
            workSheet_excel.Pictures.Add(6, 6, 20, 10, pic_path[1]);
            workSheet_excel.Pictures.Add(6, 11, 20, 16,pic_path[2]);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        private void ToExcel()
        {
            string sqlquary = string.Format(@"SELECT TOP 6 日期, 当日平均进尺数, 上部平均数, 中部平均数, 下部平均数
                                FROM 支架报表数据表  WHERE 日期<='{0}'ORDER BY 日期 DESC ", today);
            //string[] name = { "日期", "", "", "", "下部平均数" };
            SqlDataAdapter sda = new SqlDataAdapter(sqlquary, sqlConn);
            DataTable datatable = new DataTable();
            sda.Fill(datatable);
            DataRow[] dataRow = datatable.Select();
            int column = 14;
            if (dataRow.Length > 0) {

                for (int i = 0; i < dataRow.Length; i++) {

                    workSheet_excel.Cells[1, column].PutValue(dataRow[i]["当日平均进尺数"].ToString());
                    workSheet_excel.Cells[2, column].PutValue(dataRow[i]["上部平均数"].ToString());
                    workSheet_excel.Cells[3, column].PutValue(dataRow[i]["中部平均数"].ToString());
                    workSheet_excel.Cells[4, column].PutValue(dataRow[i]["下部平均数"].ToString());
                    double aver = (double.Parse(dataRow[i]["下部平均数"].ToString()) + 
                                   double.Parse(dataRow[i]["上部平均数"].ToString()) + 
                                   double.Parse(dataRow[i]["中部平均数"].ToString())) / 3;
                    workSheet_excel.Cells[5, column].PutValue(aver);

                    if(column >= 4)
                        column -= 2;

                }

            }
            Cell cellItem = workSheet_excel.Cells["R2"];
            cellItem.PutValue(string.Format("工作面1#-55#支架阻力最大值{0}MPa，平均值{1}MPa；56#-110#支架阻力最大值{2}MPa，平均值{3}MPa；111#-149#支架阻力最大值{4}MPa，平均值{5}MPa。", max_three[0], average[0], max_three[1], average[1], max_three[2], average[2]) + behaviors);



        }

        private void InsertToDB(double[] average)
        {

            //string str = @"INSERT INTO 支架报表数据表 VALUES ( '" + today + "', " + ava + ", " + average[0] + ", " + average[1] + ", " + average[2] + ")";
            string str = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                       FROM 支架报表数据表
                                                       WHERE  [日期] = '{0}')
                                         BEGIN
                                         INSERT INTO 支架报表数据表
                                         ([日期], [当日平均进尺数], [上部平均数], [中部平均数], [下部平均数], [来压情况]) 
                                         VALUES ( '{0}', {1}, {2}, {3}, {4}, N'{5}' ) 
                                         END
                                         ELSE
                                         BEGIN
                                         UPDATE 支架报表数据表
                                         SET [当日平均进尺数] = {1}, [上部平均数] = {2},
                                         [中部平均数] = {3}, [下部平均数] = {4}, [来压情况] = N'{5}'
                                         WHERE [日期] = '{0}'   END ", today, 
                                         ava, Math.Round(average[0], 2), Math.Round(average[1], 2), Math.Round(average[2], 2), behaviors);
            SqlCommand sc = new SqlCommand(str, sqlConn);
            sc.ExecuteNonQuery();
        }

        private void Calculate()
        {
            for (int i = 0; i < locations.Length; i++)
            {

                string sqlquary = string.Format(@"SELECT CQBH, ZDATETIME, ZBH, WZ, P1
                                                                  FROM ZCSJB
                                                                  RIGHT JOIN ZCYLFJWZB
                                                                  ON ZBH=YLFJBH
                                                                  WHERE ZDATETIME BETWEEN '{0} 08:00:00' AND '{1} 07:59:59'
                                                                  AND WZ='{2}' ORDER BY P1 DESC", yestoday, today, locations[i]);
                DataTable datatable = GetDataTable(sqlquary);
                DataRow[] dataRows = datatable.Select();

                if (dataRows.Length > 0) {
                    max_three[i] = double.Parse(dataRows[0]["P1"].ToString());
                    sum = 0;
                    for (int j = 0; j < dataRows.Length; j++){
                        sum += double.Parse(dataRows[j]["P1"].ToString());
                    }
                    average[i] = sum / dataRows.Length;

                }

                
            }
            InsertToDB(average);
            
        }

        private void OpenExcel()
        {
            sqlConn.Open();
            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();
            string filepath = @"..\..\modelFile\支架模板.xlsx";
            workBook = new Workbook(filepath);
            workSheet = workBook.Worksheets[0];

            if (workBook_excel.Worksheets["Sheet4"] != null)
            {

                workSheet_excel = workBook_excel.Worksheets["Sheet4"];

            }
            else
            {

                workBook_excel.Worksheets.Add("Sheet4");
                workSheet_excel = workBook_excel.Worksheets["Sheet4"];

            }
            workSheet_excel.Copy(workSheet);
            workBook_excel.Save(excelFilePath,SaveFormat.Xlsx);
        }

        private DataTable GetDataTable(string sqlquary)
        {
            sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
            DataTable datatable = new DataTable();
            sqlDataAdapter.Fill(datatable);
            return datatable;
        }



    }
}
/*

SELECT CQBH, ZDATETIME, ZBH, WZ, P1
      FROM ZCSJB
      RIGHT JOIN ZCYLFJWZB
      ON ZBH=YLFJBH
      WHERE ZDATETIME BETWEEN '{0} 08:00:00' AND '{1} 07:59:59'
      AND WZ='{2}'
*/
