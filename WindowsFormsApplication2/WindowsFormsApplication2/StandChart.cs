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
    class StandChart
    {

        private Workbook wb = new Workbook();
        private Worksheet ws;
        private Chart[] chart = new Chart[3];
        private string  path = @"..\..\modelFile\中间表.xlsx";
        string[] pic_path = { @"..\..\modelFile\chart3.png", @"..\..\modelFile\chart4.png", @"..\..\modelFile\chart5.png" };
        string[] pic_name = { "上部支架阻力曲线", "中部支架阻力曲线", "下部支架阻力曲线" };
        public StandChart() { }
        public void GetChart() {
            wb.LoadFromFile(path);
            ws = wb.Worksheets[0];
            int row = 5;
            int column = ws.Range.ColumnCount - 20;
            int lastRow = 8;
            int  lastColumn = ws.Range.ColumnCount;
            for (int i = 0; i < chart.Length; i++)
            {
                int k = 6;
                int r = row ;
                chart[i] = ws.Charts.Add();
                chart[i].ChartType = ExcelChartType.Line;
                //chart[i].DataRange = ws.Range[row-1, column, lastRow-1, lastColumn];
                //chart[i].SeriesDataFromRange = false;
                //chart 位置 A20-D30
                chart[i].LeftColumn = 1;
                chart[i].TopRow = 20;
                chart[i].RightColumn = 10;
                chart[i].BottomRow = 40;

                chart[i].ChartTitle = pic_name[i];
                chart[i].ChartTitleArea.IsBold = true;
                chart[i].ChartTitleArea.Size = 10;
                //设置横坐标的标题
                chart[i].PrimaryCategoryAxis.Title = "时间";
                chart[i].PrimaryCategoryAxis.Font.IsBold = true;
                chart[i].PrimaryCategoryAxis.HasMajorGridLines = true;
                chart[i].PrimaryCategoryAxis.TitleArea.IsBold = true;

                chart[i].PrimaryValueAxis.Title = "压力值(max)";
                chart[i].PrimaryValueAxis.HasMajorGridLines = true;
                chart[i].PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
                chart[i].PrimaryValueAxis.MinValue = 5;
                chart[i].PrimaryValueAxis.TitleArea.IsBold = true;
                //循环绘制
                for (int m = 0; m < 3; m++)
                {
                    Spire.Xls.Charts.ChartSerie cs = chart[i].Series.Add();
                    //每个series的名字
                    cs.Name = ws.Range[string.Format(@"B{0}", k++)].Value;//B6~B8
                    //设置横坐标
                    cs.CategoryLabels = ws.Range[row, column, row, lastColumn];
                    r++;
                    cs.Values = ws.Range[r, column, r, lastColumn];
                    cs.SerieType = ExcelChartType.LineMarkers;
                    cs.Format.Options.IsVaryColor = true;
                    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;
                }

                row += 5;
                lastRow += 5 ;
            }
            Image[] images = wb.SaveChartAsImage(ws);
            if (images != null)
            {
                int i = 0;
                while (i < images.Length)
                {

                    images[i].Save(pic_path[i], ImageFormat.Png);
                    i++;

                }
            }

            //wb.SaveToFile(path);
        }
    }
}
