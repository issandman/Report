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
            int row = 5, lastRow = 8;
            int  lastColumn = ws.Range.ColumnCount;
            int column = lastColumn - 21;
            for (int i = 0; i < chart.Length; i++)
            {
                chart[i] = ws.Charts.Add();
                chart[i].ChartType = ExcelChartType.Line;
                chart[i].DataRange = ws.Range[row, column, lastRow, lastColumn];
                //chart 位置 A20-D30
                chart[i].LeftColumn = 1;
                chart[i].TopRow = 20;
                chart[i].RightColumn = 6;
                chart[i].BottomRow = 30;
                chart[i].ChartTitle = pic_name[i];
                chart[i].ChartTitleArea.IsBold = true;
                chart[i].ChartTitleArea.Size = 10;
                //设置横坐标的标题
                chart[i].PrimaryCategoryAxis.Title = "时间";
                chart[i].PrimaryCategoryAxis.Font.IsBold = true;
                chart[i].PrimaryCategoryAxis.TitleArea.IsBold = true;


                //y
                chart[i].PrimaryValueAxis.Title = "压力值";
                chart[i].PrimaryValueAxis.HasMajorGridLines = false;
                chart[i].PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
                chart[i].PrimaryValueAxis.MinValue = 5;
                chart[i].PrimaryValueAxis.TitleArea.IsBold = true;
                //循环绘制
                foreach (Spire.Xls.Charts.ChartSerie cs in chart[i].Series)
                {

                    cs.Format.Options.IsVaryColor = true;
                    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

                }

                chart[i].LeftColumn += 10;
                chart[i].RightColumn += 10;
                row += 5;
                lastRow += 5;

            }
                Image[] images = wb.SaveChartAsImage(ws);
                for (int i = 0; i < chart.Length; i++)
                {
                    images[i].Save(pic_path[i], ImageFormat.Png);
                }
            wb.SaveToFile(path);
        }
    }
}
