using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;
using System.Threading;

namespace WindowsFormsApplication2
{
    public partial class Form2 : Form
    {
        BaseDateTest bdt = new BaseDateTest();
        GetTime gt = new GetTime();
        DrillingTest dt = new DrillingTest();
        VibrationTest vt = new VibrationTest();
        StandTest standtest = new StandTest();
        StandMidFile smf = new StandMidFile();
        Comprehensive ch =  new Comprehensive();
        WorkFaceWeighting wfw = new WorkFaceWeighting();
        /**********************************************************************************/
        string working_face = "2101";    //工作面
        private double[,] data = new double[21, 16];
        public List<ComboBox> list_ComboBox = new List<ComboBox>();   //不同深度煤粉量链表
        double[] data_test = new double[10];
        string type;    //钻屑法胶运辅运选择
        int tab = 0;    //钻屑法未设距工作面距离标记

        int count = 1;   //钻孔计数器
        int count1 = 0;   //单孔深度计数器
        int tag = 0;    //标志是否出现卡钻等异常情况

        string year = DateTime.Now.Year.ToString();     //日期
        string month = DateTime.Now.Month.ToString();
        string day = DateTime.Now.Day.ToString();
        string year1 = DateTime.Now.AddDays(-1).Year.ToString();     //前一天日期
        string month1 = DateTime.Now.AddDays(-1).Month.ToString();
        string day1 = DateTime.Now.AddDays(-1).Day.ToString();

        double auxiliary;   //辅运顺槽进尺
        double rubber;      //胶运顺槽进尺

        private double[,] data_footage = new double[3, 3];  //支架进尺信息
        bool[] sign = new bool[8];  //保存预览按钮点击标志
        bool sign_ = true; //同上

        public Form2()
        {
            InitializeComponent();
            loadList();
        }

        private void loadList() //钻屑法前端控件链表
        {
            list_ComboBox.Add(comboBox2);
            list_ComboBox.Add(comboBox3);
            list_ComboBox.Add(comboBox4);
            list_ComboBox.Add(comboBox5);
            list_ComboBox.Add(comboBox6);
            list_ComboBox.Add(comboBox7);
            list_ComboBox.Add(comboBox8);
            list_ComboBox.Add(comboBox9);
            list_ComboBox.Add(comboBox10);
            list_ComboBox.Add(comboBox11);
            list_ComboBox.Add(comboBox12);
            list_ComboBox.Add(comboBox13);
            list_ComboBox.Add(comboBox14);
            list_ComboBox.Add(comboBox15);
            list_ComboBox.Add(comboBox16);
        }

        private void loaddata()     //存储钻屑法数据
        {
            //double[,] data = judge_load();
            int a = 0;  //标识符（保证数据有问题时可修改不被删除）
            double max = 0;  //单孔最大值
            double max_depth = 0;   //最大值对应深度
            double add = 0;   //累加
            type = comboBox18.Text;
            //将煤粉量读入数组
            for (int i = 0; i < list_ComboBox.Count(); i++)
            {
                string temp = list_ComboBox[i].Text.ToString();
                if (String.IsNullOrWhiteSpace(temp))
                {
                    continue;
                }
                try
                {
                    if (temp == "吸钻" || temp == "卡钻" || temp == "煤炮" || temp == "卡钻吸钻")
                    {
                        switch (temp)
                        {
                            case "吸钻":
                                tag = -1; break;
                            case "卡钻":
                                tag = -2; break;
                            case "煤炮":
                                tag = -3; break;
                            case "卡钻吸钻":
                                tag = -4; break;

                        }
                        data[i, count - 1] = tag;
                    }
                    else
                    {
                        tag = 0;
                        data[i, count - 1] = double.Parse(temp);

                        add += double.Parse(temp);
                        if (max < double.Parse(temp))
                        {
                            max = double.Parse(temp);
                            max_depth = i + 1;
                        }
                        count1++;
                    }
                }
                catch (Exception ex)
                {
                    string tips = "数据的五种格式为：“煤粉量数字”、“吸钻”、“卡钻”、“卡钻吸钻”“煤炮”";
                    tips = ex.Message + tips;
                    MessageBox.Show(tips);
                    count--;
                    a = 1;
                }
            }
            //将钻孔距工作面距离、最大值、对应深度、平均值读入数组，清空内容
            if (a == 0)
            {
                try
                {
                    data[15, count - 1] = double.Parse(comboBox17.Text.ToString());
                    data[16, count - 1] = max;
                    data[17, count - 1] = max_depth;
                    data[18, count - 1] = Math.Round(add / count1, 2);
                    if (comboBox18.Text.ToString() == "辅运顺槽")
                        data[19, count - 1] = 0;
                    else if (comboBox18.Text.ToString() == "胶运顺槽")
                        data[19, count - 1] = 1;

                    comboBox17.Text = null;     //清空填空内容
                    foreach (ComboBox combobox in list_ComboBox)
                    {
                        combobox.Text = null;
                    }
                    tab = 1;
                }
                catch
                {
                    MessageBox.Show("未设置该钻孔距工作面距离");
                    count--;
                    tab = 0;
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)   //日期选择控件
        {
            string date = dateTimePicker1.Value.ToShortDateString(); //取年月日
            year = dateTimePicker1.Value.Year.ToString();
            month = dateTimePicker1.Value.Month.ToString();
            day = dateTimePicker1.Value.Day.ToString();

            year1 = dateTimePicker1.Value.AddDays(-1).Year.ToString();
            month1 = dateTimePicker1.Value.AddDays(-1).Month.ToString();
            day1 = dateTimePicker1.Value.AddDays(-1).Day.ToString();


            string date_t = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;

            gt.setToday(date_t);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);

            //MessageBox.Show(date);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //工作面选择控件
        {
            working_face = comboBox1.Text.ToString();
            //MessageBox.Show(working_face);
        }
        
        private void button9_Click(object sender, EventArgs e)  //钻屑下一个按钮
        {
            loaddata();
            count++;
            label38.Text = count.ToString() + "号钻孔";
        }

        private void button10_Click(object sender, EventArgs e)  //钻屑上一个按钮
        {
            count--;
            label38.Text = count.ToString() + "号钻孔";
        }

        private void button4_Click(object sender, EventArgs e)  //钻屑法保存预览按钮
        {
            Format format = new Format();
            string date = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;

            gt.setToday(date);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);

            string filename = string.Format("{0}工作面钻屑监测统计表{1}年{2}月.xlsx", working_face, year, month);
            string headline = string.Format("{0}月{1}日轨顺槽钻屑监测统计表", month, day);
            string headline1 = string.Format("{0}月{1}日运顺槽钻屑监测统计表", month, day);
            string path = @"F:\" + filename;

            string test = @"F:\钻屑监测统计表模板_test.xlsx";

            if (comboBox2.Text != "")
            {
                loaddata();
                count++;
                label38.Text = count.ToString() + "号钻孔";
            }


            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string constr = "server=192.168.1.111;database=UPRESSURE;uid=sa;pwd=sdkjdx";
            if (tab == 1)
            {
                using (SqlConnection sqlConnection = new SqlConnection(constr_test))
                {
                    bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                    if (conok)
                    {
                        for (int i = 0; i < count - 1; i++)
                        {
                            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[钻屑法数据表] 
                                                                 WHERE  [日期] = '{0}' AND [钻孔编号] = '{1}'AND [胶运] = '{21}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[钻屑法数据表]([日期], [钻孔编号],
                                                                                    [1m], [2m], [3m], [4m],
                                                                                    [5m], [6m], [7m], [8m],
                                                                                    [9m], [10m], [11m], [12m],
                                                                                    [13m], [14m], [15m], 
                                                                                    [距工作面距离], [单孔最大值], 
                                                                                    [最大孔深], [单孔平均], [胶运])
                                                   VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', 
                                                          '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', 
                                                          '{17}', '{18}', '{19}', '{20}', '{21}' )
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[钻屑法数据表]
                                                   SET [1m] = '{2}', [2m] = '{3}',[3m] = '{4}',[4m] = '{5}',[5m] = '{6}',
                                                       [6m] = '{7}',[7m] = '{8}',[8m] = '{9}',[9m] = '{10}',[10m] = '{11}',
                                                       [11m] = '{12}',[12m] = '{13}',[13m] = '{14}',[14m] = '{15}',[15m] = '{16}',
                                                       [距工作面距离] = '{17}', [单孔最大值] = '{18}', [最大孔深] = '{19}',
                                                       [单孔平均] = '{20}'
                                                   WHERE [日期] = '{0}' AND [钻孔编号] = '{1}' AND [胶运] = '{21}'
                                                   END",
                                                               date, i + 1, data[0, i], data[1, i], data[2, i], data[3, i],
                                                               data[4, i], data[5, i], data[6, i], data[7, i], data[8, i],
                                                               data[9, i], data[10, i], data[11, i], data[12, i], data[13, i],
                                                               data[14, i], data[15, i], data[16, i], data[17, i], data[18, i], data[19, i]);
                            SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                            cmd_ins.ExecuteNonQuery();
                        }
                        MessageBox.Show("完成");
                        count = 1;
                        label38.Text = count.ToString() + "号钻孔";
                    }
                }
            }

            try
            {
                //调用钻屑法按钮
                dt.Start(gt);
                MessageBox.Show("钻屑法ok");
            }
            catch (Exception)
            {
                MessageBox.Show("钻屑法错误 请重新检查数据");
                throw;
            }
            sign[3] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)    //钻屑法类型选择
        {
            if (MessageBox.Show("是否已经输完" + type + "数据", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (comboBox2.Text != "")
                {
                    loaddata();
                }
                //loaddata();
                string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
                //string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sdkjdx";
                string date = year + "-" + month + "-" + day;
                if (tab == 1)
                {
                    using (SqlConnection sqlConnection = new SqlConnection(constr_test))
                    {
                        bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                        if (conok)
                        {
                            for (int i = 0; i < count; i++)
                            {
                                string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[钻屑法数据表] 
                                                                 WHERE  [日期] = '{0}' AND [钻孔编号] = '{1}'AND [胶运] = '{21}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[钻屑法数据表]([日期], [钻孔编号],
                                                                                    [1m], [2m], [3m], [4m],
                                                                                    [5m], [6m], [7m], [8m],
                                                                                    [9m], [10m], [11m], [12m],
                                                                                    [13m], [14m], [15m], 
                                                                                    [距工作面距离], [单孔最大值], 
                                                                                    [最大孔深], [单孔平均], [胶运])
                                                   VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', 
                                                          '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', 
                                                          '{17}', '{18}', '{19}', '{20}', '{21}' )
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[钻屑法数据表]
                                                   SET [1m] = '{2}', [2m] = '{3}',[3m] = '{4}',[4m] = '{5}',[5m] = '{6}',
                                                       [6m] = '{7}',[7m] = '{8}',[8m] = '{9}',[9m] = '{10}',[10m] = '{11}',
                                                       [11m] = '{12}',[12m] = '{13}',[13m] = '{14}',[14m] = '{15}',[15m] = '{16}',
                                                       [距工作面距离] = '{17}', [单孔最大值] = '{18}', [最大孔深] = '{19}',
                                                       [单孔平均] = '{20}'
                                                   WHERE [日期] = '{0}' AND [钻孔编号] = '{1}' AND [胶运] = '{21}'
                                                   END",
                                                                   date, i + 1, data[0, i], data[1, i], data[2, i], data[3, i],
                                                                   data[4, i], data[5, i], data[6, i], data[7, i], data[8, i],
                                                                   data[9, i], data[10, i], data[11, i], data[12, i], data[13, i],
                                                                   data[14, i], data[15, i], data[16, i], data[17, i], data[18, i], data[19, i]);
                                SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                                cmd_ins.ExecuteNonQuery();
                            }
                        }
                    }
                    Array.Clear(data, 0, data.Length);
                    count = 1;
                    label38.Text = count.ToString() + "号钻孔";
                    MessageBox.Show("完成");
                }
            }
            else
                comboBox18.Text = type;
        }

        private void button1_Click(object sender, EventArgs e)  //基本信息保存预览
        {
            string date = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;
            string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            //插入主键时间+工作面
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[基本数据表] 
                                                                 WHERE  [日期] = '{0}' AND [工作面] = N'{1}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[基本数据表]([日期], [工作面])
                                                   VALUES('{0}', N'{1}')
                                                   END", date, working_face);
            //寻找插入数据的上一条数据
            string sqlString_find = string.Format(@"SELECT TOP 1 *
                                                    FROM [UPRESSURE].[dbo].[基本数据表]
                                                    WHERE [工作面] LIKE N'{0}'
                                                    AND [日期] < '{1}' 
                                                    ORDER BY [日期] DESC", working_face, date);


            using (SqlConnection sqlConnection = new SqlConnection(constr))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();

                    DataTable datatable = new DataTable();
                    SqlCommand cmd_find = new SqlCommand(sqlString_find, sqlConnection);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd_find))
                    {
                        da.Fill(datatable);
                    }
                    //上一条记录进尺
                    double auxiliary_y = Convert.ToDouble(datatable.Rows[0][2].ToString());
                    double rubber_y = Convert.ToDouble(datatable.Rows[0][3].ToString());
                    //总进尺
                    auxiliary = (textBox1.Text == "") ? auxiliary_y : Convert.ToDouble(textBox1.Text); //辅
                    rubber = (textBox2.Text == "") ? rubber_y : Convert.ToDouble(textBox2.Text);   //胶
                    double transport_avg = Math.Round((auxiliary + rubber) / 2, 1);
                    //当日进尺
                    double auxiliary_td = Math.Round(auxiliary - auxiliary_y, 1);
                    double rubber_td = Math.Round(rubber - rubber_y, 1);
                    double transport_td_avg = Math.Round((auxiliary_td + rubber_td) / 2, 1);
                    //涌水量
                    double water = (textBox3.Text == "") ? 0.0 : Convert.ToDouble(textBox3.Text);

                    //DataRow dr = datatable.NewRow();
                    //object[] objs = { date, working_face, auxiliary, rubber, transport_avg, auxiliary_td, rubber_td, transport_td_avg, water, textBox4.Text, textBox5.Text, textBox6.Text, (2077 - transport_avg), textBox15.Text };
                    //dr.ItemArray = objs;
                    //datatable.Rows.Add(dr);
                    //datatable写入数据库
                    string sqlString_insdata = string.Format(@"UPDATE [UPRESSURE].[dbo].[基本数据表]
                                                               SET [辅运顺槽总进尺] = '{0}', [胶运顺槽总进尺] = '{1}',
                                                                   [总进尺平均] = '{2}', [辅运当日进尺] = '{3}',
                                                                   [胶运当日进尺] = '{4}', [当日平均] = '{5}',
                                                                   [工作面涌水量] = '{6}', [初采时间] = N'{7}',
                                                                   [实测倾斜长度] = '{8}', [平均采高] = '{9}',
                                                                   [剩余推进长度] = '{10}', [时空关系] = N'{11}'
                                                               WHERE [日期] = '{12}' AND [工作面] = N'{13}'",
                                                               auxiliary, rubber, transport_avg, auxiliary_td, rubber_td, transport_td_avg, water, textBox4.Text, textBox5.Text, textBox6.Text, (2077 - transport_avg), textBox15.Text, date, working_face);
                    SqlCommand cmd_insdata = new SqlCommand(sqlString_insdata, sqlConnection);
                    cmd_insdata.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                }
            }
            //charu
            gt.setToday(date);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);

            try
            {
                bdt.Start(gt);
                MessageBox.Show("基本数据完成");
            }
            catch (Exception)
            {

                MessageBox.Show("基本数据异常");
                throw;
            }
            sign[0] = true;
            for(int i = 0; i< sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void button3_Click(object sender, EventArgs e)  //工作面来压情况保存预览
        {
            string date = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";

            //插入或更新来压情况
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[工作面来压情况] 
                                                                 WHERE  [日期] = '{0}' AND [工作面] = N'{1}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[工作面来压情况]([日期], [工作面],
                                                                                      [已来压次数],[上次位置],
                                                                                      [上次时间],[步距],
                                                                                      [本次来压情况],[持续距离],
                                                                                      [预计下次时间],[预计下次位置],
                                                                                      [预计下次步距],[下一危险区域名称],
                                                                                      [距离危险区域])
                                                   VALUES('{0}', N'{1}', N'{2}', N'{3}', N'{4}', N'{5}', N'{6}', N'{7}', N'{8}', N'{9}', N'{10}', N'{11}', N'{12}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[工作面来压情况]
                                                   SET [已来压次数] = N'{2}', [上次位置] = N'{3}',
                                                   [上次时间] = N'{4}', [步距] = N'{5}',
                                                   [本次来压情况] = N'{6}', [持续距离] = N'{7}',
                                                   [预计下次时间] = N'{8}', [预计下次位置] = N'{9}',
                                                   [预计下次步距] = N'{10}', [下一危险区域名称] = N'{11}',
                                                   [距离危险区域] = N'{12}'
                                                   WHERE [日期] = '{0}' AND [工作面] = N'{1}'
                                                   END",
                                                   date, working_face,
                                                   (textBox21.Text == "") ? null : textBox21.Text, //已来压次数
                                                   (textBox20.Text == "") ? null : textBox20.Text, //上次位置
                                                   (textBox19.Text == "") ? null : textBox19.Text, //上次时间
                                                   (textBox18.Text == "") ? null : textBox18.Text, //步距
                                                   (textBox17.Text == "") ? null : textBox17.Text, //本次来压情况
                                                   (textBox16.Text == "") ? null : textBox16.Text, //持续距离
                                                   (textBox26.Text == "") ? null : textBox26.Text, //预计下次时间
                                                   (textBox24.Text == "") ? null : textBox24.Text, //预计下次位置
                                                   (textBox22.Text == "") ? null : textBox22.Text, //预计下次步距
                                                   (textBox25.Text == "") ? null : textBox25.Text, //下一危险区域名称
                                                   (textBox23.Text == "") ? null : textBox23.Text  //距离危险区域
                                                   ); 
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    MessageBox.Show("完成");
                }
            }
            gt.setToday(date);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);
            try
            {
                wfw.Start(gt);
                MessageBox.Show("来压ok");
            }
            catch (Exception)
            {

                throw;
            }
            sign[1] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void button6_Click(object sender, EventArgs e)  //应力在线保存预览
        {
            this.backgroundWorker1.RunWorkerAsync();
            Form1 form = new Form1(this.backgroundWorker1);
            form.ShowDialog(this);
            form.Close();
        }

        private void button15_Click(object sender, EventArgs e) //微震数据分析按钮
        {
            string date_t = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;

            gt.setToday(date_t);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);

            try
            {
                vt.Start();
                button12.Enabled = true;
                button13.Enabled = true;
            }
            catch (Exception)
            {
                MessageBox.Show("微震错误");
                throw;
            }

        }

        //微震保存预览按钮
        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("微震监测ok");
            sign[4] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }
        //图片1
        private void button12_Click(object sender, EventArgs e)
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择要上传的图片";
            //图片格式
            openFileDialog.Filter = "*.png|*.png|*.jpg|*.jpg|*.bmp|*.bmp";
            //不允许多选
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                pictureBox1.Load(filePath);
            }
            vt.AddPictures(filePath, 1);
            MessageBox.Show("图片1 ok");
        }
        //图片2
        private void button13_Click(object sender, EventArgs e)
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择要上传的图片";
            //图片格式
            openFileDialog.Filter = "*.png|*.png|*.jpg|*.jpg|*.bmp|*.bmp";
            //不允许多选
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                pictureBox2.Load(filePath);
            }
            vt.AddPictures(filePath, 2);
            MessageBox.Show("图片2 ok");
        }
        //合成按钮
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                new MergeTest().MergeAllSheet(gt);
                MessageBox.Show("生成成功");
            }
            catch (Exception)
            {
                MessageBox.Show("请关闭报表");
            }
            //new MergeTest().MergeAllSheet(gt);
            //MessageBox.Show("生成成功");

        }

        private void button7_Click(object sender, EventArgs e)  //现场显现记录保存预览
        {
            string date = year + "-" + month + "-" + day;
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[现场显现记录] 
                                                                 WHERE  [日期] = '{0}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[现场显现记录]([日期], [现场记录],
                                                                                    [微震对应], [具体对应位置],
                                                                                    [现场详细描述], [分析说明])
                                                   VALUES('{0}', N'{1}', N'{2}', N'{3}', N'{4}', N'{5}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[现场显现记录]
                                                   SET [现场记录] = N'{1}', [微震对应] = N'{2}', [具体对应位置] = N'{3}',
                                                   [现场详细描述] = N'{4}', [分析说明] = N'{5}'
                                                   WHERE [日期] = '{0}'
                                                   END", date,
                                                   (textBox11.Text == "") ? null : textBox11.Text,
                                                   (textBox27.Text == "") ? null : textBox27.Text,
                                                   (textBox12.Text == "") ? null : textBox12.Text,
                                                   (textBox13.Text == "") ? null : textBox13.Text,
                                                   (textBox14.Text == "") ? null : textBox14.Text
                                                   );
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                    MessageBox.Show("完成");
                }
            }
            gt.setToday(date);
            gt.setWorkFace(working_face);
            try
            {
                new Record().Start(gt);
                MessageBox.Show("现场ok");
            }
            catch (Exception)
            {
                MessageBox.Show("记录异常");
                throw;
            }
            sign[6] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void checkCombo1_DataFill(object sender, UtilityLibrary.Combos.TreeCombo.EventArgsTreeDataFill e) //采取措施下拉菜单内容
        {
            e.BindedControl.Nodes.Add("严格执行工作面防冲专项安全技术措施");
            e.BindedControl.Nodes.Add("可以正常生产");
            e.BindedControl.Nodes.Add("降低推进速度匀速推进");
            e.BindedControl.Nodes.Add("有冲击危险降低推进速度");
            e.BindedControl.Nodes.Add("无法正常生产");
            e.BindedControl.Nodes.Add("经治理采取措施后可正常生产");
            e.BindedControl.Nodes.Add("经钻屑检验可正常生产");
        }

        private void button8_Click(object sender, EventArgs e)  //综合分析保存预览
        {
            string date = year + "-" + month + "-" + day;
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[综合分析]
                                                                 WHERE  [日期] = '{0}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[综合分析]([日期], [综合分析],
                                                                                [采取措施], [轨顺钻孔施工地点],
                                                                                [胶运钻孔施工地点], [轨顺孔深],
                                                                                [胶运孔深], [轨顺孔间距], 
                                                                                [胶运孔间距], [轨顺设计个数],
                                                                                [胶运设计个数], [轨顺当日施工个数],
                                                                                [胶运当日施工个数], [轨顺剩余个数],
                                                                                [胶运剩余个数])
                                                   VALUES('{0}', N'{1}', N'{2}', N'{3}', N'{4}', N'{5}', N'{6}', N'{7}', N'{8}', N'{9}', N'{10}', N'{11}', N'{12}', N'{13}', N'{14}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[综合分析]
                                                   SET [综合分析] = N'{1}', [采取措施] = N'{2}', [轨顺钻孔施工地点] = N'{3}',
                                                   [胶运钻孔施工地点] = N'{4}', [轨顺孔深] = N'{5}', [胶运孔深] = N'{6}',
                                                   [轨顺孔间距] = N'{7}', [胶运孔间距] = N'{8}', [轨顺设计个数] = N'{9}',
                                                   [胶运设计个数] = N'{10}', [轨顺当日施工个数] = N'{11}', [胶运当日施工个数] = N'{12}',
                                                   [轨顺剩余个数] = N'{13}', [胶运剩余个数] = N'{14}'
                                                   WHERE [日期] = '{0}'
                                                   END", date,
                                                   (textBox42.Text == "") ? null : textBox42.Text,
                                                   (checkCombo1.Value == "") ? null : checkCombo1.Value,
                                                   (textBox28.Text == "") ? null : textBox28.Text,
                                                   (textBox29.Text == "") ? null : textBox29.Text,
                                                   (textBox32.Text == "") ? null : textBox32.Text,
                                                   (textBox33.Text == "") ? null : textBox33.Text,
                                                   (textBox34.Text == "") ? null : textBox34.Text,
                                                   (textBox35.Text == "") ? null : textBox35.Text,
                                                   (textBox37.Text == "") ? null : textBox37.Text,
                                                   (textBox36.Text == "") ? null : textBox36.Text,
                                                   (textBox39.Text == "") ? null : textBox39.Text,
                                                   (textBox38.Text == "") ? null : textBox38.Text,
                                                   (textBox41.Text == "") ? null : textBox41.Text,
                                                   (textBox40.Text == "") ? null : textBox40.Text
                                                   );
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                    MessageBox.Show("完成");
                }
            }
            gt.setToday(date);
            gt.setWorkFace(working_face);
            try
            {
                ch.Start(gt);
                MessageBox.Show("ojbk!!!");

            }
            catch (Exception)
            {

                throw;
            }
            sign[7] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void button11_Click(object sender, EventArgs e) //支架阻力保存预览
        {
            data_footage[0, 0] = (textBox43.Text == "") ? 0.0 : Convert.ToDouble(textBox43.Text);
            data_footage[0, 1] = (textBox44.Text == "") ? 0.0 : Convert.ToDouble(textBox44.Text);
            data_footage[0, 2] = (textBox45.Text == "") ? 0.0 : Convert.ToDouble(textBox45.Text);
            data_footage[1, 0] = (textBox48.Text == "") ? 0.0 : Convert.ToDouble(textBox48.Text);
            data_footage[1, 1] = (textBox47.Text == "") ? 0.0 : Convert.ToDouble(textBox47.Text);
            data_footage[1, 2] = (textBox46.Text == "") ? 0.0 : Convert.ToDouble(textBox46.Text);
            data_footage[2, 0] = (textBox51.Text == "") ? 0.0 : Convert.ToDouble(textBox51.Text);
            data_footage[2, 1] = (textBox50.Text == "") ? 0.0 : Convert.ToDouble(textBox50.Text);
            data_footage[2, 2] = (textBox49.Text == "") ? 0.0 : Convert.ToDouble(textBox49.Text);

            string date = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;
            string behaviors = (textBox52.Text == "") ? null : textBox52.Text;

            gt.setToday(date);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);

            smf.Start(gt, data_footage);
            MessageBox.Show("ok1");
            new StandChart().GetChart();
            MessageBox.Show("ok2");
            standtest.Start(gt, (rubber + auxiliary) / 2, behaviors);
            MessageBox.Show("他妈的ojbk！！！！！");
            //try
            //{

            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("支架数据异常");
            //}
            sign[2] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            string date_t = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;

            gt.setToday(date_t);
            gt.setYesToday(date_y);
            gt.setWorkFace(working_face);
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[地表沉降数据] 
                                                                 WHERE  [日期] = '{0}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[地表沉降数据]([日期], [观察日期], [最大沉降量],
                                                                                    [平均沉降量], [最大沉降位置])
                                                   VALUES('{0}', N'{1}', N'{2}', N'{3}', N'{4}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[地表沉降数据]
                                                   SET [观察日期] = N'{1}', [最大沉降量] = N'{2}', 
                                                   [平均沉降量] = N'{3}', [最大沉降位置] = N'{4}'
                                                   WHERE [日期] = '{0}'
                                                   END", date_t,
                                                   (textBox7.Text == "") ? null : textBox7.Text,
                                                   (textBox8.Text == "") ? null : textBox8.Text,
                                                   (textBox9.Text == "") ? null : textBox9.Text,
                                                   (textBox10.Text == "") ? null : textBox10.Text
                                                   );
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                    //MessageBox.Show("完成");
                }
            }

            try
            {

                StressTest st = new StressTest(working_face);
                st.Start(gt);

                for (int i = 0; i < 100; i++)
                {
                    Thread.Sleep(100);
                    worker.ReportProgress(i);
                }

                MessageBox.Show("完成");
                MessageBox.Show("应力监测ok");
            }
            catch (Exception)
            {
                MessageBox.Show("应力监测数据错误");
            }
            sign[5] = true;
            for (int i = 0; i < sign.Length; i++)
            {
                if (sign[i] == false)
                    sign_ = false;
            }
            if (sign_ == true)
                button2.Enabled = true;
            sign_ = true;
        }



    }
}
