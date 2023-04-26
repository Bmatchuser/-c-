using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;


namespace 串口助手
{
    public partial class User : Form
    {
        //定义两个全局变量，接收datagridview1点击的那一行所对应的版本，项目号和汇总编号
        private string 版本号2, 项目号2, 汇总编号2,规范项目名称2;
        //定义三个double全局变量，方便“投标指标”按钮的计算
        private double 正线路基比例, 正线隧道比例, 正线桥梁比例;
        public User()
        {
            InitializeComponent();
            toolStripStatusLabel5.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            timer1.Start();
            this.textBox1.Text = "1";
            LoadProject();
        }
        //设置的计时器，用于显示当前的时间
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel5.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel5_Click(object sender, EventArgs e)
        {

        }

        
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void User_Load(object sender, EventArgs e)
        {
            this.comboBox1.SelectedIndex = 0;
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 0;
        }

        Dao2 dao2 = new Dao2();

        //从本地数据库读 版本_项目表
        private void LoadProject() 
        {

            string sql = "select * from  版本_项目表  order by 版本,项目号 desc";
            IDataReader dr = dao2.read(sql);
            while (dr.Read())
            {
                string 版本, 项目号, 汇总编号, 速度目标值, 单双线, 定额体系, 规范项目名称, 材料信息价, 地区, 省份;

                汇总编号 = dr["汇总编号"].ToString();
                版本 = dr["版本"].ToString();
                项目号 = dr["项目号"].ToString();
                速度目标值 = dr["速度目标值"].ToString();
                单双线 = dr["单双线"].ToString();
                定额体系 = dr["定额体系"].ToString();
                规范项目名称 = dr["规范项目名称"].ToString();
                材料信息价 = dr["材料信息价"].ToString();
                地区 = dr["地区"].ToString();
                省份 = dr["省份"].ToString();
                string[] str = { 版本, 项目号, 规范项目名称, 汇总编号, 速度目标值, 单双线, 定额体系, 材料信息价, 地区, 省份 };
                dataGridView1.Rows.Add(str);
            }
            dr.Close();


        }
   
        private void User_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();//结束整个程序
        }
        //返回
        private void button3_Click(object sender, EventArgs e)
        {
            login login = new login();
            login.Show();
            this.Hide();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //点击dataGridView1的某一行以后在dataGridView2显示相应的数据
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.Rows.Clear();
            //从点击的那一行获取想要的值
            string str4,sql;
            str4 = this.textBox1.Text.Trim();
            版本号2 = dataGridView1.SelectedCells[0].Value.ToString();
            项目号2 = dataGridView1.SelectedCells[1].Value.ToString();
            规范项目名称2 = dataGridView1.SelectedCells[2].Value.ToString();
            汇总编号2 = dataGridView1.SelectedCells[3].Value.ToString();
            //如果版本号为空，说明是导入的excel数据
            if (版本号2.Equals("")) {
                sql = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' order by 条目编号 asc";
                IDataReader dr = dao2.read(sql);
                while (dr.Read())
                {
                    string 条目编号,章别,节号,工程或费用项目名称,单位1,工程数量1,合价,指标1,合价7,指标15;
                    double 指标3, 指标4, str5, 合价1, 合价2;
                    
                    str5 = (float)Convert.ToSingle(str4);
                    条目编号 = dr["条目编号"].ToString();
                    章别 = dr["章别"].ToString();
                    节号 = dr["节号"].ToString();
                    工程或费用项目名称 = dr["工程或费用项目名称"].ToString();
                    单位1 = dr["单位1"].ToString();
                    工程数量1 = dr["工程数量1"].ToString();
                    合价 = dr["合价"].ToString();
                    指标1 = dr["指标1"].ToString();
                    if (合价.Contains("\\"))
                    {
                        string 合价3 = 合价.Substring(0,合价.IndexOf("\\") -1 );
                        string 合价4 = 合价.Substring(合价.IndexOf("\\") + 1);
                        double 合价5 = double.Parse(合价3) * str5;
                        double 合价6 = double.Parse(合价4) * str5;
                        合价7 = 合价5 + "\\" + 合价6;
                    }
                    else if (!合价.Equals(""))
                    {
                        string 合价s = 0 + 合价;
                        double 合价3 = double.Parse(合价s);
                        合价7 = (合价3 * str5).ToString();
                    }
                    else
                    {
                        合价7 = "";
                    }

                    if (指标1.Contains("\\"))
                    {
                        string 指标11 = 指标1.Substring(0, 指标1.IndexOf("\\") );
                        string 指标12 = 指标1.Substring(指标1.IndexOf("\\") + 1);
                        double 指标13 = double.Parse(指标11) * str5;
                        double 指标14 = double.Parse(指标12) * str5;
                        指标15 = 指标13 + "\\" + 指标14;
                    }
                    else if (!指标1.Equals(""))
                    {
                        string 指标2 = 0 + 指标1;
                        double 指标14 = double.Parse(指标2);
                        指标15 = (指标14 * str5).ToString();
                    }
                    else {
                        指标15 = "";
                    }
                    
                    


                    string[] header = { "汇总编号", "条目编号", "章别", "节号", "项目名称", "单位", "工程数量", "合价", "指标", };
                    string[] str = { 条目编号, 章别, 节号, 工程或费用项目名称, 单位1, 工程数量1, 合价7, 指标15 };
                    dataGridView2.Rows.Add(str);

                    

                    dataGridView2.Columns[0].Visible = true;
                    dataGridView2.Columns[1].Visible = true;
                    dataGridView2.Columns[2].Visible = true;
                    dataGridView2.Columns[3].Visible = true;
                    dataGridView2.Columns[4].Visible = true;
                    dataGridView2.Columns[5].Visible = true;
                    dataGridView2.Columns[6].Visible = true;
                    dataGridView2.Columns[7].Visible = true;
                    dataGridView2.Columns[8].Visible = false;
                    dataGridView2.Columns[0].HeaderCell.Value = "条目编号";
                    dataGridView2.Columns[1].HeaderCell.Value = "章别";
                    dataGridView2.Columns[2].HeaderCell.Value = "节号";
                    dataGridView2.Columns[3].HeaderCell.Value = "工程或费用项目名称";
                    dataGridView2.Columns[4].HeaderCell.Value = "单位1";
                    dataGridView2.Columns[5].HeaderCell.Value = "工程数量1";
                    dataGridView2.Columns[6].HeaderCell.Value = "合价";
                    dataGridView2.Columns[7].HeaderCell.Value = "指标1";
                    dataGridView2.ColumnHeadersVisible = true;
                }
            }
            else
            {
                sql = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' order by 条目编号 asc";
                IDataReader dr = dao2.read(sql);
                while (dr.Read())
                {
                    string 版本号, 项目号, 汇总编号, 条目编号, 速度目标值, 单双线, 定额体系, 章别, 节号, 工程或费用项目名称, 单位1, 单位2, 工程数量1, 工程数量2, 建筑工程费, 安装工程费, 设备工器具费, 其他费, 合价, 指标1, 指标2, 座数, 清单编码;
                    double 指标3, 指标4, str5, 合价1, 合价2;
                    str5 = (float)Convert.ToSingle(str4);
                    汇总编号 = dr["汇总编号"].ToString();
                    条目编号 = dr["条目编号"].ToString();
                    章别 = dr["章别"].ToString();
                    节号 = dr["节号"].ToString();
                    工程或费用项目名称 = dr["工程或费用项目名称"].ToString();
                    单位1 = dr["单位1"].ToString();
                    工程数量1 = dr["工程数量1"].ToString();

                    合价 = dr["合价"].ToString();
                    合价1 = double.Parse(合价);
                    合价2 = 合价1 * str5;

                    指标1 = dr["指标1"].ToString();
                    指标3 = double.Parse(指标1);
                    指标4 = 指标3 * str5;
                    string[] header = { "汇总编号", "条目编号", "章别", "节号", "项目名称", "单位", "工程数量", "合价", "指标", };
                    string[] str = { 汇总编号, 条目编号, 章别, 节号, 工程或费用项目名称, 单位1, 工程数量1, 合价2.ToString(), 指标4.ToString() };
                    dataGridView2.Columns[2].Visible = true;
                    dataGridView2.Columns[3].Visible = true;
                    dataGridView2.Columns[4].Visible = true;
                    dataGridView2.Columns[5].Visible = true;
                    dataGridView2.Columns[6].Visible = true;
                    dataGridView2.Columns[7].Visible = true;
                    dataGridView2.Columns[8].Visible = true;
                    dataGridView2.Columns[0].HeaderCell.Value = "汇总编号";
                    dataGridView2.Columns[1].HeaderCell.Value = "条目编号";
                    dataGridView2.Columns[2].HeaderCell.Value = "章别";
                    dataGridView2.Columns[3].HeaderCell.Value = "节号";
                    dataGridView2.Columns[4].HeaderCell.Value = "项目名称";
                    dataGridView2.Columns[5].HeaderCell.Value = "单位";
                    dataGridView2.Columns[6].HeaderCell.Value = "工程数量";
                    dataGridView2.Columns[7].HeaderCell.Value = "合价";
                    dataGridView2.Columns[8].HeaderCell.Value = "指标";
                    dataGridView2.ColumnHeadersVisible = true;

                    dataGridView2.Rows.Add(str);
                }
            
            }
            

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
        //点击搜索以后在dataGridView2显示相应的数据
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            //从三个下拉列表获取到查询需要的三个条件值
            string str1,str2,str3,str4,sql;
            str1 = this.comboBox1.SelectedItem.ToString();
            str2 = this.comboBox2.SelectedItem.ToString();
            str3 = this.comboBox3.SelectedItem.ToString();
            
            if (this.textBox1.Text.Trim().Equals(""))
            {
                this.textBox1.Text = "1";
            }

            if (this.comboBox1.SelectedIndex == 0 && this.comboBox2.SelectedIndex == 0 && this.comboBox3.SelectedIndex == 0) 
            {
                sql = "select * from 版本_项目表 where 速度目标值 = '" + "" + "' and 单双线='" + "" + "' and 定额体系= '" + ""+ "'";
            }
            else if (this.comboBox1.SelectedIndex != 0 && this.comboBox2.SelectedIndex == 0 && this.comboBox3.SelectedIndex == 0)
            {
                sql = "select * from 版本_项目表 where 速度目标值 = '" + str1 + "' ";
                Console.Out.WriteLine(sql);
            }
            else if (this.comboBox1.SelectedIndex == 0 && this.comboBox2.SelectedIndex != 0 && this.comboBox3.SelectedIndex == 0)
            {
                sql = "select * from 版本_项目表 where  单双线='" + str2 + "' ";
            }
            else if (this.comboBox1.SelectedIndex == 0 && this.comboBox2.SelectedIndex == 0 && this.comboBox3.SelectedIndex != 0)
            {
                sql = "select * from 版本_项目表 where  定额体系= '" + str3 + "'";
            }
            else if (this.comboBox1.SelectedIndex != 0 && this.comboBox2.SelectedIndex != 0 && this.comboBox3.SelectedIndex == 0)
            {
                sql = "select * from 版本_项目表 where 速度目标值 = '" + str1 + "' and 单双线='" + str2 + "' ";
            }
            else if (this.comboBox1.SelectedIndex != 0 && this.comboBox2.SelectedIndex == 0 && this.comboBox3.SelectedIndex != 0)
            {
                sql = "select * from 版本_项目表 where 速度目标值 = '" + str1 + "' and 定额体系= '" + str3 + "'";
            }
            else if (this.comboBox1.SelectedIndex == 0 && this.comboBox2.SelectedIndex != 0 && this.comboBox3.SelectedIndex != 0)
            {
                sql = "select * from 版本_项目表 where  单双线='" + str2 + "' and 定额体系= '" + str3 + "'";
            }
            else 
            {
                sql = "select * from 版本_项目表 where 速度目标值 = '" + str1 + "' and 单双线='" + str2 + "' and 定额体系= '" + str3 + "'";
                Console.Out.WriteLine(sql);
            }
           
            
            IDataReader dr = dao2.read(sql);
            while (dr.Read())
            {
                string 版本, 项目号, 汇总编号, 速度目标值, 单双线, 定额体系,规范项目名称,材料信息价,地区,省份;
               
                
                汇总编号 = dr["汇总编号"].ToString();
                版本 = dr["版本"].ToString();
                项目号 = dr["项目号"].ToString();
                速度目标值 = dr["速度目标值"].ToString();
                单双线 = dr["单双线"].ToString();
                定额体系 = dr["定额体系"].ToString();
                规范项目名称 = dr["规范项目名称"].ToString();
                材料信息价 = dr["材料信息价"].ToString();
                地区 = dr["地区"].ToString();
                省份 = dr["省份"].ToString();
                string[] str = { 版本, 项目号, 规范项目名称, 汇总编号, 速度目标值, 单双线, 定额体系, 材料信息价, 地区, 省份 };
                dataGridView1.Rows.Add(str);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void excel_out_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);// 添加工作薄
            excel.Visible = true;
            // 生成字段名称
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                excel.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText;
            }
            // 填充数据
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    if (dataGridView2[j, i].Value == null)
                    {
                        dataGridView2[j, i].Value = "";
                    }
                    else
                    {
                        if (dataGridView2[j, i].ValueType == typeof(string))
                            excel.Cells[i + 2, j + 1] = "'" + dataGridView2[j, i].Value.ToString();
                        else
                            excel.Cells[i + 2, j + 1] = dataGridView2[j, i].Value.ToString();
                    }
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            版本号2 = dataGridView1.SelectedCells[0].Value.ToString();
            规范项目名称2 = dataGridView1.SelectedCells[2].Value.ToString();
            string sql;
            dataGridView2.Rows.Clear();

            if (版本号2.Equals("")){
                sql = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2  + "' and  章别 <>'' or 节号 <>'' or 工程或费用项目名称 = '第一部分' or 工程或费用项目名称 = '以上各章合计'or 工程或费用项目名称 = '以上总计'or 工程或费用项目名称 = '第二部分' or 工程或费用项目名称 = '第三部分' or 工程或费用项目名称 = '第四部分'or 工程或费用项目名称 = '估算总额' order by 条目编号 asc";
            }
            else {
                sql = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and  章别 <>'' or 节号 <>'' or 工程或费用项目名称 = '第一部分' or 工程或费用项目名称 = '以上各章合计'or 工程或费用项目名称 = '以上总计'or 工程或费用项目名称 = '第二部分' or 工程或费用项目名称 = '第三部分' or 工程或费用项目名称 = '第四部分'or 工程或费用项目名称 = '估算总额' order by 条目编号 asc";
            }
            IDataReader dr = dao2.read(sql);
            while (dr.Read())
            {
                string 章别,节号,工程或费用项目名称,单位1,工程数量1,合价,指标1;
                章别 = dr["章别"].ToString();
                节号 = dr["节号"].ToString();
                工程或费用项目名称 = dr["工程或费用项目名称"].ToString();
                单位1 = dr["单位1"].ToString();
                工程数量1 = dr["工程数量1"].ToString();
                合价 = dr["合价"].ToString();
                指标1 = dr["指标1"].ToString();

                string[] str = { 章别, 节号, 工程或费用项目名称, 单位1, 工程数量1, 合价, 指标1 };
                dataGridView2.ColumnHeadersVisible = true;
                dataGridView2.Columns[0].Visible = true;
                dataGridView2.Columns[1].Visible = true;
                dataGridView2.Columns[2].Visible = true;
                dataGridView2.Columns[3].Visible = true;
                dataGridView2.Columns[4].Visible = true;
                dataGridView2.Columns[5].Visible = true;
                dataGridView2.Columns[6].Visible = true;
                dataGridView2.Columns[7].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[0].HeaderCell.Value = "章别";
                dataGridView2.Columns[1].HeaderCell.Value = "节号";
                dataGridView2.Columns[2].HeaderCell.Value = "工程或费用项目名称";
                dataGridView2.Columns[3].HeaderCell.Value = "单位1";
                dataGridView2.Columns[4].HeaderCell.Value = "工程数量";
                dataGridView2.Columns[5].HeaderCell.Value = "合价";
                dataGridView2.Columns[6].HeaderCell.Value = "指标";

                dataGridView2.Rows.Add(str);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            string sql, sql4, sql5, sql6, sql7, sql8, sql9, sql10, sql11, sql12, sql13, sql14, sql15, sql16, sql17, sql18, sql19, sql20, sql21 ;
            版本号2 = dataGridView1.SelectedCells[0].Value.ToString();
            规范项目名称2 = dataGridView1.SelectedCells[2].Value.ToString();
            dataGridView2.Rows.Clear();

            if (版本号2.Equals(""))
            {
                sql = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' ";
            }
            else
            {
                sql = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' ";
            }
            IDataReader dr = dao2.read(sql);
            if (dr.Read()) {
                //第一行：项目名称
                string[] str = { "项目名称", dr["规范项目名称"].ToString() };
                dataGridView2.Rows.Add(str);

                //第二行：设计阶段
                 string[] str1 = {"设计阶段", dr["设计阶段"].ToString()};
                 dataGridView2.Rows.Add(str1);

                 //第三行：标准
                 string[] str2 = { "标准", dr["速度目标值"].ToString() };
                 dataGridView2.Rows.Add(str2);
            }
            //第四行：正线长度(km)
            if (版本号2.Equals(""))
            {
                sql4 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' and 单位1 like '" + "%正线公里%" + "' and  章别 <>''";
            }
            else
            {
                sql4 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 单位1 = '" + "正线公里" + "' and  章别 <>''";
            }
            IDataReader dr4 = dao2.read(sql4);
            if (dr4.Read())
            {
                string 正线公里 = dr4["工程数量1"].ToString();
                if (正线公里.Contains("\\"))
                {
                    string 正线公里1 = 正线公里.Substring(正线公里.IndexOf("\\") + 1);
                    string[] str = { "正线长度（km）", 正线公里1 };
                    dataGridView2.Rows.Add(str);
                }
                else
                {
                    string[] str2 = { "正线长度（km）", dr4["工程数量1"].ToString() };
                    dataGridView2.Rows.Add(str2);
                }
               
            }
            //第五行：正线路基比例(%)  路基公里/正线公里（单位1）
            if (版本号2.Equals(""))
            {
                sql5 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' and 单位1 like '" + "%路基公里%" + "' and  章别 <>''";
            }
            else
            {
                sql5 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 单位1 = '" + "路基公里" + "' and  章别 <>''";
            }
            IDataReader dr5 = dao2.read(sql5);
            if (dr5.Read())
            {
                double 路基公里2, 正线公里1;
                string 路基公里 = dr5["工程数量1"].ToString();
                string 正线公里 = dr4["工程数量1"].ToString();
                if (路基公里.Contains("\\"))
                {
                    string 路基公里1 = 路基公里.Substring(路基公里.IndexOf("\\") + 1);
                    路基公里2 = double.Parse(路基公里1);
                    正线公里1 = double.Parse(正线公里);
                    正线路基比例 = 路基公里2 / 正线公里1;
                    string[] str = { "正线路基比例（%）", 正线路基比例.ToString() };
                    dataGridView2.Rows.Add(str);
                }
                else {
                    路基公里2 = double.Parse(路基公里);
                    正线公里1 = double.Parse(正线公里);
                    正线路基比例 = 路基公里2 / 正线公里1;
                    string[] str2 = { "正线路基比例（%）", 正线路基比例.ToString() };
                    dataGridView2.Rows.Add(str2);
                }
                
                
            }
            //第六行：正线隧道比例(%)  隧道长度/正线公里
            if (版本号2.Equals(""))
            {
                sql6 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' and 单位1 like '" + "%隧道公里%" + "' and  章别 <>''";
            }
            else
            {
                sql6 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 单位1 = '" + "隧道公里" + "' and  章别 <>''";
            }
            IDataReader dr6 = dao2.read(sql6);
            if (dr6.Read())
            {
                double 隧道公里2, 正线公里1;
                string 隧道公里 = dr6["工程数量1"].ToString();
                string 正线公里 = dr4["工程数量1"].ToString();

                if (隧道公里.Contains("\\"))
                {
                    string 隧道公里1 = 隧道公里.Substring(隧道公里.IndexOf("\\") + 1);
                    隧道公里2 = double.Parse(隧道公里1);
                    正线公里1 = double.Parse(正线公里);
                    正线隧道比例 = 隧道公里2 / 正线公里1;
                    string[] str = { "正线隧道比例（%）", 正线隧道比例.ToString() };
                    dataGridView2.Rows.Add(str);
                }
                else
                {
                    隧道公里2 = double.Parse(隧道公里);
                    正线公里1 = double.Parse(正线公里);
                    正线隧道比例 = 隧道公里2 / 正线公里1;
                    string[] str2 = { "正线隧道比例（%）", 正线隧道比例.ToString() };
                    dataGridView2.Rows.Add(str2);
                }
            }
            //第七行：正线桥梁比例(%)  1-路基比例-隧道比例
                正线桥梁比例 = 1 - 正线路基比例 - 正线隧道比例;
                string[] str3 = { "正线桥梁比例（%）", 正线桥梁比例.ToString() };
                dataGridView2.Rows.Add(str3);
            //第八行：正线桥隧比例(%)  桥梁比例+隧道比例
                double 正线桥隧比例 = 正线桥梁比例+正线隧道比例;
                string[] str4 = { "正线桥隧比例（%）", 正线桥隧比例.ToString() };
                dataGridView2.Rows.Add(str4);
            //第九行：人工价格
            if (版本号2.Equals(""))
            {
                sql9 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' ";
            }
            else
            {
                sql9 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' ";
            }
            IDataReader dr9 = dao2.read(sql9);
            if (dr9.Read())
            {
                string[] str = { "人工价格", dr["人工单价"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十行：编制期价格
            if (版本号2.Equals(""))
            {
                sql10 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'";
            }
            else
            {
                sql10 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' ";
            }
            IDataReader dr10 = dao2.read(sql10);
            if (dr10.Read())
            {
                string[] str = { "编制期价格", dr10["材料信息价"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十一行：拆迁及征地费用（万元/正线公里）
            string content = "一";
            if (版本号2.Equals(""))
            {
                sql11 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content + "'";
            }
            else
            {
                sql11 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content + "' ";
            }
            IDataReader dr11 = dao2.read(sql11);
            if (dr11.Read())
            {
                string[] str = { "拆迁及征地费用（万元/正线公里）", dr11["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十二行：路基(万元/路基公里)
            string content2 = "二";
            if (版本号2.Equals(""))
            {
                sql12 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content2 + "'";
            }
            else
            {
                sql12 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content2 + "' ";
            }
            IDataReader dr12 = dao2.read(sql12);
            if (dr12.Read())
            {
                string[] str = { "路基(万元/路基公里)", dr12["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十三行：桥涵（万元/桥梁公里）
            string content3 = "三";
            if (版本号2.Equals(""))
            {
                sql13 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content3 + "'";
            }
            else
            {
                sql13 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content3 + "' ";
            } IDataReader dr13 = dao2.read(sql13);
            if (dr13.Read())
            {
                string[] str = { "桥涵（万元/桥梁公里）", dr13["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十四行：隧道及明洞（万元/隧道公里）
            string content4 = "四";
            if (版本号2.Equals(""))
            {
                sql14 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' and  章别 = '" + content4 + "'";
            }
            else
            {
                sql14 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content4 + "' ";
            } IDataReader dr14 = dao2.read(sql14);
            if (dr14.Read())
            {
                string[] str = { "隧道及明洞（万元/隧道公里）", dr14["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十五行：轨道（万元/正线公里）
            string content5 = "五";
            if (版本号2.Equals(""))
            {
                sql15 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content5 + "'";
            }
            else
            {
                sql15 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content5 + "' ";
            } IDataReader dr15 = dao2.read(sql15);
            if (dr15.Read())
            {
                string[] str = { "轨道（万元/正线公里）", dr15["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十六行：四电（万元/正线公里）
            string content6 = "六";
            string content7 = "七";
            if (版本号2.Equals(""))
            {
                sql16 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content6 + "'";
            }
            else
            {
                sql16 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content6 + "' ";
            }
            if (版本号2.Equals(""))
            {
                sql17 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content7 + "'";
            }
            else
            {
                sql17 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content7 + "' ";
            } IDataReader dr16 = dao2.read(sql16);
            IDataReader dr17 = dao2.read(sql17);
            if (dr16.Read() && dr17.Read())
            {
                double 第六章指标1, 第七章指标1,指标;
                string 第六章指标 = dr16["指标1"].ToString();
                string 第七章指标 = dr17["指标1"].ToString();
                第六章指标1 = double.Parse(第六章指标);
                第七章指标1 = double.Parse(第七章指标);
                指标 = 第六章指标1 + 第七章指标1;
                string[] str = { "四电（万元/正线公里）", 指标.ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十七行：房屋（万元/正线公里）
            string content8 = "八";
            if (版本号2.Equals(""))
            {
                sql18 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content8 + "'";
            }
            else
            {
                sql18 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content8 + "' ";
            } IDataReader dr18 = dao2.read(sql18);
            if (dr18.Read())
            {
                string[] str = { "房屋（万元/正线公里）", dr18["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十八行：其他运营生产设备及建筑物（万元/正线公里）
            string content9 = "九";
            if (版本号2.Equals(""))
            {
                sql19 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content9 + "'";
            }
            else
            {
                sql19 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content9 + "' ";
            } IDataReader dr19 = dao2.read(sql19);
            if (dr19.Read())
            {
                string[] str = { "其他运营生产设备及建筑物（万元/正线公里）", dr19["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第十九行：大型临时设施和过渡工程（万元/正线公里）
            string content10 = "十";
            if (版本号2.Equals(""))
            {
                sql20 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "'  and  章别 = '" + content10 + "'";
            }
            else
            {
                sql20 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 章别 = '" + content10 + "' ";
            } IDataReader dr20 = dao2.read(sql20);
            if (dr20.Read())
            {
                string[] str = { "大型临时设施和过渡工程（万元/正线公里）", dr20["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            //第二十行：静态投资技术经济指标
            string content11 = "%静态投资%";
            if (版本号2.Equals(""))
            {
                sql21 = "select * from 综合概算汇总表2 where 规范项目名称= '" + 规范项目名称2 + "' and 工程或费用项目名称 like '" + content11 + "' ";
            }
            else
            {
                sql21 = "select * from 综合概算汇总表 where 版本号= '" + 版本号2 + "' and 项目号= '" + 项目号2 + "'  and 汇总编号= '" + 汇总编号2 + "' and 工程或费用项目名称 like '" + content11 + "' ";
            } IDataReader dr21 = dao2.read(sql21);
            if (dr21.Read())
            {
                string[] str = { "静态投资技术经济指标", dr21["指标1"].ToString() };
                dataGridView2.Rows.Add(str);
            }
            dataGridView2.Columns[0].HeaderCell.Value = "项目";
            dataGridView2.Columns[1].HeaderCell.Value = "内容";
            dataGridView2.Columns[0].Visible = true;
            dataGridView2.Columns[1].Visible = true;
            dataGridView2.Columns[2].Visible = false;
            dataGridView2.Columns[3].Visible = false;
            dataGridView2.Columns[4].Visible = false;
            dataGridView2.Columns[5].Visible = false;
            dataGridView2.Columns[6].Visible = false;
            dataGridView2.Columns[7].Visible = false;
            dataGridView2.Columns[8].Visible = false;
            //隐藏表头
            dataGridView2.ColumnHeadersVisible = false;
        }

        private void button1_Click_2(object sender, EventArgs e)
        {

        }
    }
}
