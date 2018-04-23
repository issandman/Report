using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class ChartFuck
    {
        Workbook wb = new Workbook();
        Worksheet ws;
        private string path = @"..\..\modelFile\中间表.xlsx";
        string[] pic_path = { @"..\..\modelFile\chart3.png", @"..\..\modelFile\chart4.png", @"..\..\modelFile\chart5.png" };
        string[] pic_name = { "上部支架阻力曲线", "中部支架阻力曲线", "下部支架阻力曲线" };
        Chart chart;
        public ChartFuck() {
            wb.LoadFromFile(path);
            ws = wb.Worksheets[0];
            int row = 5;
            int column = ws.Range.ColumnCount - 20;
            int lastRow = 8;
            int lastColumn = ws.Range.ColumnCount;
            chart = ws.Charts.Add();
            chart.ChartType = ExcelChartType.Line;
            //chart.DataRange = ws.Range[row, column, lastRow, lastColumn];
            chart.SeriesDataFromRange = false;
            chart.LeftColumn = 1;
            chart.TopRow = 20;
            chart.RightColumn = 10;
            chart.BottomRow = 40;
            chart.ChartTitle = pic_name[0];
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 10;


            Spire.Xls.Charts.ChartSerie cs = chart.Series.Add();
            cs.Name = ws.Range["B6"].Value;//B6~B8
            cs.CategoryLabels = ws.Range[5, 10, 5, 11];
            cs.Values = ws.Range[6, 10, 6, 11];
            cs.DataFormat.ShowActiveValue = true;
            cs.Format.Options.IsVaryColor = true;
            cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;

            Spire.Xls.Charts.ChartSerie cs2 = chart.Series.Add();
            cs2.Name = ws.Range["B7"].Value;//B6~B8
            cs2.CategoryLabels = ws.Range[5, 10, 5, 11];
            cs2.Values = ws.Range[7, 10, 7, 11];
            cs2.DataFormat.ShowActiveValue = true;
            cs2.Format.Options.IsVaryColor = true;
            cs2.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;

            //int r = row;
            //int k = 6;
            //for (int m = 0; m < 3; m++)
            //{
            //    Spire.Xls.Charts.ChartSerie cs = chart.Series.Add();
            //    //每个series的名字
            //    cs.Name = ws.Range[string.Format(@"B{0}", k++)].Value;//B6~B8
            //    //设置横坐标
            //    cs.CategoryLabels = ws.Range[row, column, row, lastColumn];

            //    r++;
            //    //cs.Values = ws.Range[r, column, r, lastColumn];
            //    cs.DataFormat.ShowActiveValue = true;
            //    cs.Format.Options.IsVaryColor = true;
            //    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = false;
            //}

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

            wb.SaveToFile(path);

        }
    }
}
