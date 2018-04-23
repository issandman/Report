using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class StandMidFile
    {
        private Workbook wb;
        private Worksheet ws;
        private Cells cells;
        private string today, yestoday, workface;
        private double[,] ava = new double[3,3];
        private string[] times = { "早班", "中班", "夜班"};
        private string[] locations = { "上部", "中部", "下部" };
        private static string constr = "server=.;database=ylkdb;uid=sa;pwd=sakjdx";
        StandChart sc = new StandChart();
        private string path = @"..\..\modelFile\中间表.xlsx";
        //数据库连接
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;
        int max = 0;

        public StandMidFile() {
        }
        public void Start(GetTime gt, double[,] ava) {

            this.ava = ava;
            this.today = gt.getDateToday();
            this.yestoday = gt.getDateYestoday();
            this.workface = gt.getWorkFace();

            openexcel();
            ToExcel();
            //sc.GetChart();
            wb.Save(path, SaveFormat.Xlsx);

        }
        private void ToExcel()
        {
            //今日上中下部的所有值
            string sqlquary = string.Format(@"SELECT ZDATETIME, ZBH, WZ, P1
                                      FROM ZCSJB
                                      RIGHT JOIN ZCYLFJWZB
                                      ON ZBH=YLFJBH
                                      WHERE ZDATETIME BETWEEN    '{0} 08:00:00'  AND  '{1} 07:59:59'   ", yestoday, today);
            string sqlquary_2 = "SELECT YLFJBH, DYZJBH, WZ FROM ZCYLFJWZB";
            DataTable datatable = GetDataTable(sqlquary);
            DataTable datatable_2 = GetDataTable(sqlquary_2);
            max = cells.MaxColumn + 1;
            string date;
            for (int i = 0, j = 0 ; i < times.Length; j++, i++)
            {
                if (i == times.Length - 1) {
                    date = today;
                }
                else {
                    date = yestoday;
                }
                //全部平均 某班某部所有值的平均值
                //最大值   某班某部所有值的最大一个
                //最大平均：某班某部每一个支架编号有一个最大值 的平均
                cells[1, max].PutValue(date);
                cells[2, max].PutValue(times[i]);

                //上部 5 6 7
                cells[3, max].PutValue(ava[i,j]);
                cells[4, max].PutValue(date + times[i] + ava[i, j] + "米");
                DBToExcel("上部", times[i], datatable, datatable_2);
                //中部 10 11 12
                cells[8, max].PutValue(ava[i, j]);
                cells[9, max].PutValue(date + times[i] + ava[i, j] + "米");
                DBToExcel("中部", times[i], datatable, datatable_2);
                //下部 15 16 17
                cells[13, max].PutValue(ava[i, j]);
                cells[14, max].PutValue(date + times[i] + ava[i, j] + "米");
                DBToExcel("下部", times[i], datatable, datatable_2);

                max = cells.MaxColumn + 1;
            }
            try
            {
                wb.Save(path, SaveFormat.Xlsx);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("中间表已打开，请先关闭中间表");
            }
        }

        private void openexcel()
        {
           
            wb = new Workbook(path);
            ws = wb.Worksheets[0];
            cells = ws.Cells;
        }

        private void DBToExcel(string location, string time, DataTable dt, DataTable dt_2)
        {

            int local_m = 0, local_ava = 0, local_m_ava = 0 ;
            double sum = 0;
            double sum_everyone = 0;
            int[][] number = new int[3][];
            int l = 0;
            string select = null;
            ////////////////////////////////////////
            for (int i = 0; i < 3; i++) {
                DataRow[] dr = dt_2.Select(string.Format("WZ = '{0}'", locations[i]));
                if (dr.Length > 0) {
                    number[i] = new int[dr.Length];
                    for (int j = 0; j < number[i].Length; j++) {
                        number[i][j] = int.Parse(dr[j]["YLFJBH"].ToString());
                    }

                }

            }

            if (time.Equals("夜班")){
                select = string.Format("ZDATETIME > '{0} 00:00:00' AND ZDATETIME < '{1} 07:59:59' AND WZ='{2}'  ", today, today, location);
            }else if (time.Equals("早班")){
                select = string.Format("ZDATETIME > '{0} 08:00:00' AND ZDATETIME < '{1} 15:59:59' AND WZ='{2}'  ", yestoday, yestoday, location);
            }else if (time.Equals("中班")){
                select = string.Format("ZDATETIME > '{0} 16:00:00' AND ZDATETIME < '{1} 23:59:59' AND WZ='{2}'  ", yestoday, yestoday, location);
            }

            if (location.Equals("上部")){
                l = 0;  local_m = 7;    local_ava = 5;  local_m_ava = 6;
            }else if (location.Equals("中部")){
                l = 1;  local_m = 12;   local_ava = 10; local_m_ava = 11;
            }else if (location.Equals("下部")) {
                l = 2;  local_m = 17;   local_ava = 15; local_m_ava = 16;
            }

            DataRow[] dataRow = dt.Select(select, "P1 desc");

            if (dataRow.Length > 0) {

                cells[local_m, max].PutValue(double.Parse(dataRow[0]["P1"].ToString()));//double.Parse(dataRows_m[0]["P1"].ToString())
                sum = 0;
                for (int i = 0; i < dataRow.Length; i++)
                {

                    sum += double.Parse(dataRow[i]["P1"].ToString());

                }
                cells[local_ava, max].PutValue(sum / dataRow.Length);
                sum_everyone = 0;
                for (int i = 0; i < number[l].Length; i++)
                {

                    DataRow[] dr = dt.Select(select + string.Format("AND ZBH={0}", number[l][i]), "P1 desc");
                    if(dr.Length > 0)
                        sum_everyone += double.Parse(dr[0]["P1"].ToString());
                }
                cells[local_m_ava, max].PutValue(sum_everyone / number[l].Length);

            }


        }

        private DataTable GetDataTable(string str) {
            DataTable dt = new DataTable();
            sqlDataAdapter = new SqlDataAdapter(str, sqlConnection);
            sqlDataAdapter.Fill(dt);
            return dt;
        }
    }
}
