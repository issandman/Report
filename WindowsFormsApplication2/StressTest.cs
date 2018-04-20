using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
/*

SELECT CreatTime
      ,Channel
      ,MPId
      ,Location
      ,MsgForewarn.Value
      ,InitialValue
      ,TunnelName
      ,Depth
  FROM MsgForewarn
  INNER JOIN DataInfo
  ON Channel=MPId AND CreatTime=AcquisitionTime
  LEFT JOIN MeasurePoint
  ON Id=Channel
  WHERE CreatTime BETWEEN '2018-03-19 16:00:00' AND '2018-03-20 16:00:00'
  ORDER BY CreatTime
 


*/
namespace WindowsFormsApplication2
{
    class StressTest
    {
        //工作面
        string work_face;
        //模板变量
        private Workbook workBook;
        private Worksheet workSheet;
        //报表变量
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //报表路径
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        //连接数据库
        static string constr = "server=.;database=20171113;uid=sa;pwd=sakjdx";
        //查询唯一的channel 和 查询value、location
        string sqlquary_channel;
        string sqlquary_valus;
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;
        string today, yestoday,workface;

        //构造函数
        public StressTest(string work)
        {
            work_face = work;//
        }

        public void Start(GetTime gt)
        {

            OpenExcel();
            today = gt.getDateToday();
            yestoday = gt.getDateYestoday();
            workface = gt.getWorkFace();
            sqlquary_channel = string.Format(@"SELECT DISTINCT MPId, Depth, TunnelName FROM DataInfo 
                                                  JOIN MeasurePoint ON MPId=Id 
                                                  WHERE AcquisitionTime BETWEEN '{0} 16:00:00'AND '{1} 16:00:00' ", yestoday, today);
            sqlquary_valus = string.Format(@"SELECT CreatTime, Channel, MPId, Location, DataInfo.Value, InitialValue, TunnelName, Depth
                                                FROM MsgForewarn 
                                                INNER JOIN DataInfo ON Channel = MPId AND CreatTime = AcquisitionTime
                                                LEFT JOIN MeasurePoint ON Id=Channel   
                                                WHERE CreatTime BETWEEN '{0} 16:00:00' AND  '{1} 16:00:00' AND Location LIKE '{2}%'", yestoday, today, workface);
            string pi7 = sqlquary_channel + " AND Depth=7 AND TunnelName LIKE '皮%'";
            string pi12 = sqlquary_channel + " AND Depth=12 AND TunnelName LIKE '皮%'";
            string gui7 = sqlquary_channel + " AND Depth=7 AND TunnelName LIKE '轨%'";
            string gui12 = sqlquary_channel + " AND Depth=12 AND TunnelName LIKE '轨%'";

            string[] sql_quary = new string[] { pi7, pi12, gui7, gui12 };
            for (int i = 0; i < sql_quary.Length; i++) {

                //查询后返回 唯一的channel 组成的datatable
                DataTable dt = GetDataTable(sql_quary[i]);
                //计算
                Calculate(dt);
                //double[] number = Calculate(ways[i],deeps[i],dt);

            }//end ToExcel(...)
            Analysis();
            AddPictures();
            Subsidence();
            SaveReportFile();
        }

        private void Subsidence()
        {
            string conn = "server = .;database = UPRESSURE;uid = sa;pwd = sakjdx";
            string str = string.Format(@"SELECT * FROM 地表沉降数据 WHERE 日期='{0}'", today);
            SqlConnection sqlConnection = new SqlConnection(conn);
            SqlDataAdapter sqlDataAdapter2 = new SqlDataAdapter(str, sqlConnection);
            DataTable dt = new DataTable();
            sqlDataAdapter2.Fill(dt);
            DataRow dr = dt.Select()[0];
            Cell cellItem1 = workSheet_excel.Cells["P2"];
            Cell cellItem2 = workSheet_excel.Cells["S2"];
            Cell cellItem3 = workSheet_excel.Cells["S3"];
            Cell cellItem4 = workSheet_excel.Cells["Q4"];
            cellItem1.PutValue(dr["观察日期"].ToString());
            cellItem2.PutValue(dr["最大沉降量"].ToString());
            cellItem3.PutValue(dr["平均沉降量"].ToString());
            cellItem4.PutValue(dr["最大沉降位置"].ToString());
        }

        private void AddPictures()
        {
            CreateChart cc = new CreateChart();
            cc.GetChart(today, yestoday);
            string pic_jiaoyun_path = cc.reJiaoyunPath();
            workSheet_excel.Pictures.Add(8, 1, 20, 10, pic_jiaoyun_path);
            string pic_fuyun_path = cc.reFuyunPath();
            workSheet_excel.Pictures.Add(8, 11, 20, 19, pic_fuyun_path);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        private void Calculate(DataTable dt)
        {
            double max, location = 0;
            double max_up = 0, max_down = 0, location_up = 0, location_down = 0;
            double accum_max_up = 0, accum_max_down = 0, location_accum_up = 0, location_accum_down = 0;//累计最大 最小

            DataRow[] dataRows_channel = dt.Select();

            int deep = int.Parse(dataRows_channel[0]["Depth"].ToString());
            string way = dataRows_channel[0]["TunnelName"].ToString();

            max = CalculateMax(deep, way, ref location);
            //每一个channel
            for (int i = 0; i < dataRows_channel.Length; i++) {

                string sqlquary = string.Format(sqlquary_valus + @" AND Channel={0}", dataRows_channel[i]["MPId"] );

                DataTable datatable = GetDataTable(sqlquary);
                //同一个channel的不同时间的value
                DataRow[] dataRows_value = datatable.Select();
                //表的最旧和最新的两条
                int top = 0, bottom = dataRows_value.Length - 1;
                //当日变化值
                double value = double.Parse(dataRows_value[bottom]["Value"].ToString()) - double.Parse(dataRows_value[top]["Value"].ToString());
                if (max_up < value) {

                     max_up = value;
                     location_up = reLocation(deep, dataRows_value[bottom]["Location"].ToString());

                }

                if (max_down > value) {

                    max_down = value;
                    location_down = reLocation(deep, dataRows_value[bottom]["Location"].ToString());

                }
                //累计变化值
                double accum_value = double.Parse(dataRows_value[bottom]["Value"].ToString()) - double.Parse(dataRows_value[bottom]["InitialValue"].ToString());
                if (accum_max_up < accum_value) {

                    accum_max_up = accum_value;
                    location_accum_up = reLocation(deep, dataRows_value[bottom]["Location"].ToString());

                }
                if (accum_max_down > accum_value) {

                    accum_max_down = accum_value;
                    location_accum_down = reLocation(deep, dataRows_value[bottom]["Location"].ToString());

                }
            }// end for

            ToExcel(way, deep, max_up, max_down, max, accum_max_up, accum_max_down, 
                    location_up, location_down, location, location_accum_up,location_accum_down);
        }

        private double CalculateMax(int deep, string way,ref double location)
        {
            double max = 0;
            //???????????????????? top1
            string sqlquary = string.Format(sqlquary_valus + @" AND Depth={0} AND TunnelName LIKE '{1}' ORDER BY DataInfo.Value DESC", deep, way);
            DataTable dt = GetDataTable(sqlquary);
            max = double.Parse(dt.Select()[0]["Value"].ToString());
            location = reLocation(deep, dt.Select()[0]["Location"].ToString());

            return max;
        }

        private DataTable GetDataTable(string sqlquary)
        {
            sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
            DataTable datatable = new DataTable();
            sqlDataAdapter.Fill(datatable);
            return datatable; 
        }

        /*
            正则匹配返回监测点距离煤壁距离
        */
        private double reLocation(int deep, string str)
        {
            string rex = string.Format(@"{0}gzm(轨道巷|皮带顺槽)(.+)m处{1}m深", work_face, deep);

            if (Regex.IsMatch(str, rex))
            {

                 return double.Parse(Regex.Match(str, rex).Groups[2].Value);

            }

            return 0;

        }

        /*
        [ForewarnLevel]  =  0 ：正常
                            1 ：-5
                            2 ：黄色指标
                            3 ：红色指标

            分析结论
            */
        public void Analysis() {
            int red = 0, yellow = 0;
            Cell cellItem = workSheet_excel.Cells["N2"];
            string str = string.Format(@"SELECT ForewarnLevel FROM MsgForewarn WHERE CreatTime BETWEEN '{0} 16:00:00' 
                                                                AND '{1} 16:00:00'
                                                                AND ForewarnLevel != 1",
                                                              yestoday,
                                                                today);
            //存储所有的ForewarnLevel
            DataTable datatable = GetDataTable(str);
            foreach (DataRow dr in datatable.Select()) {

                if (int.Parse(dr["ForewarnLevel"].ToString()) == 2)
                    yellow++;
                if (int.Parse(dr["ForewarnLevel"].ToString()) == 3)
                    red++;

            }

            if (yellow == 1 || yellow == 0 || red == 0){

                cellItem.PutValue("有1个黄色指标，无冲击危险");

            }
            else if (red > 1){  
                 
                cellItem.PutValue( red + "个红色指标，红色预警  有冲击危险" );

            }
            else if (yellow > 1 || red == 1) {

                cellItem.PutValue( yellow + "个黄色指标" + red + "个红色指标，黄色预警  有冲击危险");

            }

        }

        /*
            写入excel
        */
        private void ToExcel(string way, int deep, 
                                double up, double down, double max, double accum_up, double accum_down, 
                                double location_up, double location_down, double location,
                                double location_accum_up, double location_accum_down)
        {
            Cell cellItem1 = workSheet_excel.Cells["B3"];
            Cell cellItem2 = workSheet_excel.Cells["B5"];
            cellItem1.PutValue(work_face + "胶运顺槽");
            cellItem2.PutValue(work_face + "辅运顺槽");


            if (way.Equals("皮带顺槽") && deep == 7) {

                Cell cellItem_max = workSheet_excel.Cells["D3"];
                Cell cellItem_up = workSheet_excel.Cells["F3"];
                Cell cellItem_down = workSheet_excel.Cells["H3"];
                Cell cellItem_location = workSheet_excel.Cells["E3"];
                Cell cellItem_location_up = workSheet_excel.Cells["G3"];
                Cell cellItem_location_down = workSheet_excel.Cells["I3"];
                Cell cellItem_accum_up = workSheet_excel.Cells["J3"];
                Cell cellItem_accum_down = workSheet_excel.Cells["L3"];
                Cell cellItem_location_accum_up = workSheet_excel.Cells["K3"];
                Cell cellItem_location_accum_down = workSheet_excel.Cells["M3"];

                cellItem_max.PutValue(max);
                cellItem_location.PutValue(location);

                if(up != 0)
                    cellItem_up.PutValue(up);
                if(down != 0)
                    cellItem_down.PutValue(down);
                if (location_down != 0)
                    cellItem_location_down.PutValue(location_down);
                if (location_up != 0)
                    cellItem_location_up.PutValue(location_up);

                if (accum_up != 0) {
                    cellItem_accum_up.PutValue(accum_up);
                }
                if (accum_down != 0) {
                    cellItem_accum_down.PutValue(accum_down);
                }
                if (location_accum_up != 0) {
                    cellItem_location_accum_up.PutValue(location_accum_up);
                }
                if (location_accum_down != 0) {
                    cellItem_location_accum_down.PutValue(location_accum_down);
                }

                
            }else if (way.Equals("皮带顺槽") && deep == 12)
            {

                Cell cellItem_max = workSheet_excel.Cells["D4"];
                Cell cellItem_up = workSheet_excel.Cells["F4"];
                Cell cellItem_down = workSheet_excel.Cells["H4"];
                Cell cellItem_location = workSheet_excel.Cells["E4"];
                Cell cellItem_location_up = workSheet_excel.Cells["G4"];
                Cell cellItem_location_down = workSheet_excel.Cells["I4"];

                Cell cellItem_accum_up = workSheet_excel.Cells["J4"];
                Cell cellItem_accum_down = workSheet_excel.Cells["L4"];
                Cell cellItem_location_accum_up = workSheet_excel.Cells["K4"];
                Cell cellItem_location_accum_down = workSheet_excel.Cells["M4"];

                cellItem_max.PutValue(max);
                cellItem_location.PutValue(location);

                if (up != 0)
                    cellItem_up.PutValue(up);
                if (down != 0)
                    cellItem_down.PutValue(down);
                if (location_down != 0)
                    cellItem_location_down.PutValue(location_down );
                if (location_up != 0)
                    cellItem_location_up.PutValue(location_up );
                if (accum_up != 0)
                {
                    cellItem_accum_up.PutValue(accum_up);
                }
                if (accum_down != 0)
                {
                    cellItem_accum_down.PutValue(accum_down);
                }
                if (location_accum_up != 0)
                {
                    cellItem_location_accum_up.PutValue(location_accum_up);
                }
                if (location_accum_down != 0)
                {
                    cellItem_location_accum_down.PutValue(location_accum_down);
                }
            }
            else if (way.Equals("轨道巷") && deep == 7)
            {

                Cell cellItem_max = workSheet_excel.Cells["D5"];
                Cell cellItem_up = workSheet_excel.Cells["F5"];
                Cell cellItem_down = workSheet_excel.Cells["H5"];
                Cell cellItem_location = workSheet_excel.Cells["E5"];
                Cell cellItem_location_up = workSheet_excel.Cells["G5"];
                Cell cellItem_location_down = workSheet_excel.Cells["I5"];

                Cell cellItem_accum_up = workSheet_excel.Cells["J5"];
                Cell cellItem_accum_down = workSheet_excel.Cells["L5"];
                Cell cellItem_location_accum_up = workSheet_excel.Cells["K5"];
                Cell cellItem_location_accum_down = workSheet_excel.Cells["M5"];

                cellItem_max.PutValue(max);
                cellItem_location.PutValue(location );

                if (up != 0)
                    cellItem_up.PutValue(up);
                if (down != 0)
                    cellItem_down.PutValue(down);
                if (location_down != 0)
                    cellItem_location_down.PutValue(location_down );
                if (location_up != 0)
                    cellItem_location_up.PutValue(location_up);

                if (accum_up != 0)
                {
                    cellItem_accum_up.PutValue(accum_up);
                }
                if (accum_down != 0)
                {
                    cellItem_accum_down.PutValue(accum_down);
                }
                if (location_accum_up != 0)
                {
                    cellItem_location_accum_up.PutValue(location_accum_up);
                }
                if (location_accum_down != 0)
                {
                    cellItem_location_accum_down.PutValue(location_accum_down);
                }
            }
            else if (way.Equals("轨道巷") && deep == 12)
            {

                Cell cellItem_max = workSheet_excel.Cells["D6"];
                Cell cellItem_up = workSheet_excel.Cells["F6"];
                Cell cellItem_down = workSheet_excel.Cells["H6"];
                Cell cellItem_location = workSheet_excel.Cells["E6"];
                Cell cellItem_location_up = workSheet_excel.Cells["G6"];
                Cell cellItem_location_down = workSheet_excel.Cells["I6"];

                Cell cellItem_accum_up = workSheet_excel.Cells["J6"];
                Cell cellItem_accum_down = workSheet_excel.Cells["L6"];
                Cell cellItem_location_accum_up = workSheet_excel.Cells["K6"];
                Cell cellItem_location_accum_down = workSheet_excel.Cells["M6"];

                cellItem_max.PutValue(max);
                cellItem_location.PutValue(location );

                if (up != 0)
                    cellItem_up.PutValue(up);
                if (down != 0)
                    cellItem_down.PutValue(down);
                if (location_down != 0)
                    cellItem_location_down.PutValue(location_down);
                if (location_up != 0)
                    cellItem_location_up.PutValue(location_up);

                if (accum_up != 0)
                {
                    cellItem_accum_up.PutValue(accum_up);
                }
                if (accum_down != 0)
                {
                    cellItem_accum_down.PutValue(accum_down);
                }
                if (location_accum_up != 0)
                {
                    cellItem_location_accum_up.PutValue(location_accum_up);
                }
                if (location_accum_down != 0)
                {
                    cellItem_location_accum_down.PutValue(location_accum_down);
                }
            }

        }

        private void SaveReportFile()
        {
            //  设置执行公式计算 - 如果代码中用到公式，需要设置计算公式，导出的报表中，公式才会自动计算
            workBook_excel.CalculateFormula(true);
            //  保存文件
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        private void OpenExcel()
        {
            string FilePath = @"..\..\modelFile\应力监测模板.xlsx";
            workBook = new Workbook(FilePath);
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            if (workBook_excel.Worksheets["Sheet6"] != null)
                workSheet_excel = workBook_excel.Worksheets["Sheet6"];
            else
            {
                workBook_excel.Worksheets.Add("Sheet6");
                workSheet_excel = workBook_excel.Worksheets["Sheet6"];
            }

            workSheet_excel.Copy(workSheet);
        }

        /*
            创造图表            
        */
    }
}
