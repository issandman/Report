
using CCWin.SkinClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Spire.Xls;
using System.Drawing.Imaging;

namespace WindowsFormsApplication2
{
    /*
    SELECT CONVERT(varchar(100), AcquisitionTime, 8) AS Time
	  ,AcquisitionTime
      ,MPId
      ,DataInfo.Value
      ,MsgForewarn.Memo
  FROM DataInfo
  JOIN MsgForewarn
  ON MPId=Channel AND AcquisitionTime=CreatTime
  WHERE CONVERT(varchar(100), AcquisitionTime, 8) LIKE '%:00:%'
  AND AcquisitionTime BETWEEN '{0}' AND '{1}' AND MsgForewarn.Memo='{3}'
    */
    class CreateChart
    {
        //连接数据库
        private static string constr = "server=.;database=20171113;uid=sa;pwd=sakjdx";
        
        private SqlConnection sqlConnection = new SqlConnection(constr);
        private Spire.Xls.Workbook wb = new Spire.Xls.Workbook();
        private Spire.Xls.Worksheet ws;
        private Spire.Xls.Workbook wb2 = new Spire.Xls.Workbook();
        private Spire.Xls.Worksheet ws2;
        private Spire.Xls.Chart chart;
        private static int chart_len = 6;//6个项目
        //string[] names_1 = new string[chart_len];// = new string[] { "皮3", "皮4", "皮5", "皮6", "皮7", "皮8" };
        string[] name_1 = new string[chart_len];//轨 MsgForewarn.Memo LIKE '轨%'";
        string[] name_2 = new string[chart_len];//皮 MsgForewarn.Memo LIKE '皮%'";
        private string pic_jiaoyun_path = @"..\..\modelFile\chart1.png";
        private string pic_fuyun_path = @"..\..\modelFile\chart2.png";
        public CreateChart() {
        }
        public string reJiaoyunPath() {
            return pic_jiaoyun_path;
        }
        public string reFuyunPath() {
            return pic_fuyun_path;
        }

        public void GetChart(string today,string yestoday) {
            string sqlquary = string.Format(@"SELECT CONVERT(varchar(100), AcquisitionTime, 8) AS Time
                                                  ,MPId
                                                  ,DataInfo.Value
                                                  ,MsgForewarn.Memo
                                              FROM DataInfo
                                              JOIN MsgForewarn ON MPId=Channel AND AcquisitionTime=CreatTime
                                              WHERE AcquisitionTime BETWEEN '{0} 16:00:00' AND '{1} 16:00:00' ", yestoday, today);
            GetName(sqlquary);
            PutValue(name_1, "胶运顺槽", sqlquary);
            PutValue2(name_2, "辅运顺槽", sqlquary);
            //wb.SaveToFile(@"C:\Users\14439\Desktop\yingpanhao\报表\图表aaaaa.xlsx", ExcelVersion.Version2013);

        }
        private void GetName(string sqlquary)
            {

                string str = sqlquary + @" AND MsgForewarn.Memo LIKE '皮%'";
                string str_2 = sqlquary +@" AND MsgForewarn.Memo LIKE '轨%'";

                DataTable datatable = GetDataTable(str);
                DataRow[] dataRows = datatable.Select();
                for (int i = 0; i < name_1.Length; i++) {

                    name_1[i] = dataRows[i]["Memo"].ToString();

                }

                DataTable datatable_2 = GetDataTable(str_2);
                DataRow[] dataRows_2 = datatable_2.Select();
                for (int i = 0; i < name_2.Length; i++)
                {

                    name_2[i] = dataRows_2[i]["Memo"].ToString();

                }
            }

        private void PutValue(string[] name, string title_name,string sqlquary)
                {
                    ws = wb.Worksheets[0];
                    for (int i = 1; i <= chart_len; i++)
                    {
                        //A2 2 1 行i+1列1
                        ws.Range[i + 1, 1].Text = name[i - 1];

                        string str = string.Format(sqlquary + @"AND CONVERT(varchar(100), AcquisitionTime, 8) LIKE '%:00:%' 
                                                                AND MsgForewarn.Memo = '{0}'", name[i - 1]);
                        DataTable datatable = GetDataTable(str);
                        DataRow[] dataRows = datatable.Select();
                        for (int j = 0, k = 0; k < dataRows.Length; j++, k = k + 2)
                        {
                            //B1 C1 D1
                            ws.Range[1, j + 2].Text = dataRows[k]["Time"].ToString();
                            ws.Range[i + 1, j + 2].NumberValue = Math.Round(double.Parse(dataRows[k]["Value"].ToString()), 2);
                        }
                    }
                    ChartDraw(title_name);
                }
        private void ChartDraw(string title_name)
        {
            chart = ws.Charts.Add();
            chart.ChartType = ExcelChartType.Line;
            chart.DataRange = ws.Range[1,1,7,ws.Range.ColumnCount];
//            chart.DataRange = ws.Range["A1:I7"];
            chart.LeftColumn = 2;
            chart.TopRow = 7;
            chart.RightColumn = 11;
            chart.BottomRow = 22;
            //标题名称
            chart.ChartTitle = title_name;
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;
            //设置横坐标的标题
            chart.PrimaryCategoryAxis.Title = "时间";
            chart.PrimaryCategoryAxis.Font.IsBold = true;
            chart.PrimaryCategoryAxis.TitleArea.IsBold = true;


            //y
            chart.PrimaryValueAxis.Title = "数值";
            chart.PrimaryValueAxis.HasMajorGridLines = false;
            chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
            chart.PrimaryValueAxis.MinValue = 0;
            chart.PrimaryValueAxis.TitleArea.IsBold = true;
            //循环绘制
            foreach (Spire.Xls.Charts.ChartSerie cs in chart.Series)
            {

                cs.Format.Options.IsVaryColor = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;

            }

            Image[] image = wb.SaveChartAsImage(ws);
            image[0].Save(pic_jiaoyun_path, ImageFormat.Png);
            //Image[] image = wb.SaveChartAsImage(ws);

        }

        private void PutValue2(string[] name, string title_name, string sqlquary)
        {
            ws2 = wb2.Worksheets[0];
            for (int i = 1; i <= chart_len; i++)
            {
                //A2 2 1 行i+1列1
                ws2.Range[i + 1, 1].Text = name[i - 1];

                string str = string.Format(sqlquary + @"AND CONVERT(varchar(100), AcquisitionTime, 8) LIKE '%:00:%' 
                                                        AND MsgForewarn.Memo = '{0}'", name[i - 1]);
                DataTable datatable = GetDataTable(str);
                DataRow[] dataRows = datatable.Select();
                for (int j = 0, k = 0; k < dataRows.Length; j++, k = k + 2)
                {
                    //B1 C1 D1
                    ws2.Range[1, j + 2].Text = dataRows[k]["Time"].ToString();
                    ws2.Range[i + 1, j + 2].NumberValue = Math.Round(double.Parse(dataRows[k]["Value"].ToString()), 2);
                }
            }

            ChartDraw2(title_name);
        }        
        private void ChartDraw2(string title_name)
            {
                chart = ws2.Charts.Add();
                chart.ChartType = ExcelChartType.Line;
                chart.DataRange = ws.Range[1, 1, 7, ws.Range.ColumnCount];
                chart.LeftColumn = 12;
                chart.TopRow = 7;
                chart.RightColumn = 20;
                chart.BottomRow = 22;
                //标题名称
                chart.ChartTitle = title_name;
                chart.ChartTitleArea.IsBold = true;
                chart.ChartTitleArea.Size = 12;
                //设置横坐标的标题
                chart.PrimaryCategoryAxis.Title = "时间";
                chart.PrimaryCategoryAxis.Font.IsBold = true;
                chart.PrimaryCategoryAxis.TitleArea.IsBold = true;


                //y
                chart.PrimaryValueAxis.Title = "数值";
                chart.PrimaryValueAxis.HasMajorGridLines = false;
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
                chart.PrimaryValueAxis.MinValue = 0;
                chart.PrimaryValueAxis.TitleArea.IsBold = true;
                //循环绘制
                foreach (Spire.Xls.Charts.ChartSerie cs in chart.Series)
                {

                    cs.Format.Options.IsVaryColor = true;
                    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;

                }
                Image[] image = wb2.SaveChartAsImage(ws2);
                image[0].Save(pic_fuyun_path, ImageFormat.Png);

            }

        private DataTable GetDataTable(string sqlquary)
                {
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
                    DataTable datatable = new DataTable();
                    sqlDataAdapter.Fill(datatable);
                    return datatable;
                }
        
    }
}
