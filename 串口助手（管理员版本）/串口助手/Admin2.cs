using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 串口助手
{
    public partial class Admin2 : Form
    {
        //设置两个全局变量，用于接收从admin1传送过来的数据
        private string str;  //接收项目号
        private string strs;//接收建设项目名称
        private string strss;//接收版本号
        private int i,j,k;
        public Admin2()
        {
            InitializeComponent();

        }
        //接收admin1传送过来的两个值
        public Admin2(string achieve1,string achieve2,string achieve3)
        {
    
            InitializeComponent();
            str = achieve1;//从admin1那里获取项目号
            strs = achieve2;//从admin1那里获取建设项目名称
            strss = achieve3;//从admin1那里获取版本号
        }

        private void Admin2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();//结束整个程序
        }

        //打开以后默认触发的事件
        private void Admin2_Load(object sender, EventArgs e)
        {
            //从选中的那一行中获取到要访问的项目数据库名
            Admin1 admin1 = new Admin1();
            //写入项目默认名称
            label6.Text = strs;
            //从Test方法中活的txt已存储的字符串并进行拼接
            Test test = new Test();
            string[] str2 = { "", "" };
            str2 = test.ReadText2("项目选择.txt");
            //将字符串进行拼接，拼接成想要的IP
            string str3 = str2[0] + str + str2[1];

            Console.WriteLine(str3+"测试"+str2[1]);
            //传递地址访问数据库
            Dao3 dao3 = new Dao3();
            string sql = "select * from 汇总索引表";
            DataTable dt = dao3.GetTable(sql, str3);
            this.comboBox1.DataSource = dt;
            this.comboBox1.ValueMember = "汇总编号";
            this.comboBox1.DisplayMember = "汇总编号";

            //设置默认值
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 0;
            this.comboBox4.SelectedIndex = 0;
            this.comboBox5.SelectedIndex = 0;
            this.comboBox6.SelectedIndex = 0;
            this.comboBox7.SelectedIndex = 0;
            this.comboBox8.SelectedIndex = 0;
            this.comboBox9.SelectedIndex = 0;
            textBox1.Text = "项目名";



        }

        Dao2 dao2 = new Dao2();
        Dao3 dao3 = new Dao3();

        //comboBox1被点击以后触发事件，在dataGridView1显示相应的数据
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            //从Test方法中活的txt已存储的字符串并进行拼接
            Test test = new Test();
            string[] str2 = { "", "" };
            str2 = test.ReadText2("项目选择.txt");
            //将字符串进行拼接，拼接成想要的IP
            string str3 = str2[0] + str + str2[1];

            //从项目数据库中查到对应的数据
            //通过版本号作为条件查询到对应的版本ID
            string sql2 = "select * from 汇总索引表 where 汇总编号 = '" + this.comboBox1.Text + "'";
            IDataReader dr2 = dao3.read(sql2, str3);
            while (dr2.Read())
            {
                string 汇总编号 = dr2["汇总编号"].ToString();
                //将IP存入txt文件
                test.WriteText("汇总编号选择.txt", 汇总编号);
            }
            dr2.Close();
            
            string str4 = test.ReadText("汇总编号选择.txt");

            //从项目数据库，综合概算汇总表中获得想要的数据
            string sql = "select * from 综合概算汇总表 where 汇总编号 = '" + str4 + "'";
            IDataReader dr = dao3.read(sql,str3);
            while (dr.Read())
            {
                string 汇总编号,条目编号,章别,节号,工程或费用项目名称,单位1,单位2,工程数量1,工程数量2,建筑工程费,安装工程费,设备工器具费,其他费,合价,指标1,指标2,座数,清单编码;
                汇总编号 = dr["汇总编号"].ToString();
                条目编号 = dr["条目编号"].ToString();
                章别 = dr["章别"].ToString();
                节号 = dr["节号"].ToString();
                工程或费用项目名称  = dr["工程或费用项目名称"].ToString();
                单位1 = dr["单位1"].ToString();
                工程数量1 = dr["工程数量1"].ToString();
                合价 = dr["合价"].ToString();
                指标1 = dr["指标1"].ToString();
                string[] str5 = { 汇总编号,条目编号, 章别, 节号, 工程或费用项目名称, 单位1, 工程数量1,  合价, 指标1 };
                dataGridView1.Rows.Add(str5);
            }
        }
        //返回
        private void button1_Click(object sender, EventArgs e)
        {
            Admin1 admin1 = new Admin1();

            admin1.Width = 1200;
            admin1.Height = 700;
            int height = System.Windows.Forms.SystemInformation.WorkingArea.Height;
            int width = System.Windows.Forms.SystemInformation.WorkingArea.Width;
            int formheight = admin1.Size.Height;
            int formwidth = admin1.Size.Width;
            int newformx = width / 2 - formwidth / 2;
            int newformy = height / 2 - formheight / 2;
            admin1.SetDesktopLocation(newformx, newformy);

            admin1.Show();
            this.Hide();
        }
        //如果textBox1不为空，则触发添加功能
        private void button2_Click(object sender, EventArgs e)
        {
            if(add2())
            {
                add();
            }
        }

        //如果textBox1不为空，则返回true
        public bool add2() 
        { 
            if(textBox1.Text==""||textBox2.Text=="")
            {
                MessageBox.Show("请填入数据", "提示");
                return false;
            }
            string sql = "select * from 综合概算汇总表 where 版本号= '" + strs + "' and 项目号 = '" + str + "' and 汇总编号 ='" + this.comboBox1.Text + "' ";
            IDataReader dr = dao2.read(sql);
            if (dr.Read())
            {
                MessageBox.Show("该项目已经添加，请勿重复添加！", "提示");
                return false;
            }
            return true;
        }
        //添加功能
        public bool add()
        {
            //1.从服务器获取想要的数据
            string 汇总编号 = this.dataGridView1.SelectedCells[0].Value.ToString();
            string sql = "select * from 综合概算汇总表 where 汇总编号 = '" + 汇总编号 + "'";
            //从Test方法中活的txt已存储的字符串并进行拼接
            Test test = new Test();
            string[] str2 = { "", "" };
            str2 = test.ReadText2("项目选择.txt");
            //将字符串进行拼接，拼接成想要的IP
            string str3 = str2[0] + str + str2[1];
            IDataReader dr = dao3.read(sql, str3);
            //获取要手动添加的数据
            string 速度目标值, 单双线, 定额体系, 规范项目名称,材料信息价,地区,省份,项目类型,人工单价,设计阶段;
            速度目标值 = this.comboBox2.SelectedItem.ToString();
            单双线 = this.comboBox3.SelectedItem.ToString();
            定额体系 = this.comboBox4.SelectedItem.ToString();
            地区 = this.comboBox5.SelectedItem.ToString();
            省份 = this.comboBox6.SelectedItem.ToString();
            规范项目名称 = textBox1.Text.Trim();
            材料信息价 = textBox2.Text.Trim();
            项目类型 = comboBox7.SelectedItem.ToString();
            人工单价 = comboBox8.SelectedItem.ToString();
            设计阶段 = comboBox9.SelectedItem.ToString();
            //2.将数据写入到本地数据库
            while (dr.Read())
            {
                string 版本号, 项目号, 汇总编号2, 条目编号, 章别, 节号, 工程或费用项目名称, 单位1, 单位2, 工程数量1, 工程数量2, 建筑工程费, 安装工程费, 设备工器具费, 其他费, 合价, 指标1, 指标2, 座数, 清单编码;
                版本号 = strss;
                项目号 = str;
                汇总编号2 = dr["汇总编号"].ToString();
                条目编号 = dr["条目编号"].ToString();
                章别 = dr["章别"].ToString();
                节号 = dr["节号"].ToString();
                工程或费用项目名称 = dr["工程或费用项目名称"].ToString();
                单位1 = dr["单位1"].ToString();
                单位2 = dr["单位2"].ToString();
                工程数量1 = dr["工程数量1"].ToString();
                工程数量2 = dr["工程数量2"].ToString();
                建筑工程费 = dr["建筑工程费"].ToString();
                安装工程费 = dr["安装工程费"].ToString();
                设备工器具费 = dr["设备工器具费"].ToString();
                其他费 = dr["其他费"].ToString();
                合价 = dr["合价"].ToString();
                指标1 = dr["指标1"].ToString();
                指标2 = dr["指标2"].ToString();
                座数 = dr["座数"].ToString();
                清单编码 = dr["清单编码"].ToString();
                //如果值为空默认赋值为null
                if (章别.Equals(""))
                {
                    章别 = null;
                }
                string sql2 = "insert into 综合概算汇总表 values ('" + 版本号 + "','" + 项目号 + "','" + 规范项目名称 + "','" + 汇总编号2 + "','" + 条目编号 + "','" + 速度目标值 + "','" + 单双线 + "','" + 定额体系 + "','" + 章别 + "','" + 节号 + "','" + 工程或费用项目名称 + "','" + 单位1 + "','" + 单位2 + "','" + 工程数量1 + "','" + 工程数量2 + "','" + 建筑工程费 + "','" + 安装工程费 + "','" + 设备工器具费 + "','" + 其他费 + "','" + 合价 + "','" + 指标1 + "','" + 指标2 + "','" + 座数 + "','" + 清单编码 + "','" + 材料信息价 + "','" + 地区 + "','" + 省份 + "','" + 项目类型 + "','" + 人工单价 + "','" + 设计阶段 + "')";
                i = dao2.Execute(sql2);
                
            }
            //先查询一下这个版本号，项目号是否存在，不存在的话添加，存在就不添加
            string sql3 = "select * from 版本_项目表 where 版本= '" + strs + "' and 项目号 = '" + str + "' and 汇总编号 = '" + str + "'";
            IDataReader dr2 = dao2.read(sql3);
            if (!dr2.Read())
            {

                string sql4 = "insert into 版本_项目表 values ('" + strss + "','" + str + "','" + 规范项目名称 + "','" + 速度目标值 + "','" + 单双线 + "','" + 定额体系 + "','" + 材料信息价 + "','" + 地区 + "','" + 省份 + "','" + 汇总编号 + "','" + 项目类型 + "','" + 人工单价 + "','" + 设计阶段 + "')";
                j = dao2.Execute(sql4);
               
            }

            if (i >= 0)
            {
                MessageBox.Show("添加成功", "提示");
                return true;
            }
            return false;
        }
        private void button3_Click(object sender, EventArgs e)
        {

          
        } 
        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        

        private void label3_Click(object sender, EventArgs e)
        {

       }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text == "华北")
            
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("北京");
                comboBox6.Items.Add("天津");
                comboBox6.Items.Add("河北");
                comboBox6.Items.Add("内蒙古");
                comboBox6.Items.Add("山西");
                comboBox6.Items.Add("山东");
            }
            else if (comboBox5.Text == "华东")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("上海");
                comboBox6.Items.Add("浙江");
                comboBox6.Items.Add("江苏");
                comboBox6.Items.Add("安徽");              
            }
            else if (comboBox5.Text == "华南")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("广东");
                comboBox6.Items.Add("福建");
                comboBox6.Items.Add("广西");
                comboBox6.Items.Add("海南");
            }
            else if (comboBox5.Text == "华中")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("河南");
                comboBox6.Items.Add("湖北");
                comboBox6.Items.Add("湖南");
                comboBox6.Items.Add("江西");
            }
            else if (comboBox5.Text == "西南")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("四川");
                comboBox6.Items.Add("重庆");
                comboBox6.Items.Add("昆明");
                comboBox6.Items.Add("贵州");
            }
            else if (comboBox5.Text == "东北")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("辽宁");
                comboBox6.Items.Add("吉林");
                comboBox6.Items.Add("黑龙江");
            }
            else if (comboBox5.Text == "西北")
            {
                //清除原有的内容
                comboBox6.SelectedIndex = -1;
                comboBox6.Items.Clear();
                comboBox6.Text = "";

                comboBox6.Items.Add("陕西");
                comboBox6.Items.Add("宁夏");
                comboBox6.Items.Add("青海");
                comboBox6.Items.Add("甘肃");
                comboBox6.Items.Add("新疆");
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if(add3()){
                add4();
            }
        }
        public bool add3()
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("请填入数据", "提示");
                return false;
            }
            string 规范项目名称;
            规范项目名称 = textBox1.Text.Trim();
            string sql = "select * from 综合概算汇总表 where 规范项目名称= '" + 规范项目名称 + "'  ";
            IDataReader dr = dao2.read(sql);
            if (dr.Read())
            {
                MessageBox.Show("该项目已经添加，请勿重复添加！", "提示");
                return false;
            }
            return true;
        }

        public bool add4()
        {
            //获取要手动添加的数据
            string 速度目标值, 单双线, 定额体系, 规范项目名称, 材料信息价, 地区, 省份, 项目类型, 人工单价, 设计阶段;
            速度目标值 = this.comboBox2.SelectedItem.ToString();
            单双线 = this.comboBox3.SelectedItem.ToString();
            定额体系 = this.comboBox4.SelectedItem.ToString();
            地区 = this.comboBox5.SelectedItem.ToString();
            省份 = this.comboBox6.SelectedItem.ToString();
            规范项目名称 = textBox1.Text.Trim();
            材料信息价 = textBox2.Text.Trim();
            项目类型 = comboBox7.SelectedItem.ToString();
            人工单价 = comboBox8.SelectedItem.ToString();
            设计阶段 = comboBox9.SelectedItem.ToString();
            //获取excel导入的数据
            OpenFileDialog openFile = new OpenFileDialog();
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFile.FileName;
                DataTable excelDt = ExcelHelper.ReadFromExcel(filePath);
                int count = excelDt.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    string 条目编号,章别, 节号, 工程或费用项目名称, 单位1, 工程数量1, 合价, 指标1;
                    double 工程数量2, 合价2, 指标2;
                    条目编号 = excelDt.Rows[i][0].ToString();
                    章别 = excelDt.Rows[i][1].ToString();
                    节号 = excelDt.Rows[i][2].ToString();
                    工程或费用项目名称 = excelDt.Rows[i][3].ToString();
                    单位1 = excelDt.Rows[i][4].ToString();
                    工程数量1 = excelDt.Rows[i][5].ToString();
                    合价 = excelDt.Rows[i][6].ToString();
                    指标1 = excelDt.Rows[i][7].ToString();

                   
                    //工程数量2 = double.Parse(工程数量1);
                    //合价2 = double.Parse(合价);
                    //指标2 = double.Parse(指标1);
                    if (章别.Equals(""))
                    {
                        章别 = null;
                    }
                    if (节号.Equals(""))
                    {
                        节号 = null;
                    }
                    if (工程或费用项目名称.Equals(""))
                    {
                        工程或费用项目名称 = null;
                    }
                    if (单位1.Equals(""))
                    {
                        单位1 = null;
                    }
                    if (工程数量1.Equals(""))
                    {
                        工程数量1 = null;
                    }
                    if (合价.Equals(""))
                    {
                        合价 = null;
                    }
                    if (指标1.Equals(""))
                    {
                        指标1 = null;
                    }
                    string sql = "insert into 综合概算汇总表2 (规范项目名称,条目编号,速度目标值,单双线,定额体系,章别,节号,工程或费用项目名称,单位1,工程数量1,合价,指标1,材料信息价,地区,省份,项目类型,人工单价,设计阶段)  values ('" + 规范项目名称 + "','" + 条目编号 + "','" + 速度目标值 + "','" + 单双线 + "','" + 定额体系 + "','" + 章别 + "','" + 节号 + "','" + 工程或费用项目名称 + "','" + 单位1 + "','" + 工程数量1+ "','" + 合价 + "','" + 指标1 + "','" + 材料信息价 + "','" + 地区 + "','" + 省份 + "','" + 项目类型 + "','" + 人工单价 + "','" + 设计阶段 + "')";
                    k = dao2.Execute(sql);
                }
                //把Excel读取到DataTable里面 然后再把DataTable存入数据库

            }
            //先查询一下这个版本号，项目号是否存在，不存在的话添加，存在就不添加
            string sql2 = "select * from 版本_项目表 where 规范项目名称= '" + 规范项目名称 + "' "; 
            IDataReader dr2 = dao2.read(sql2);
            if (!dr2.Read())
            {
                string sql4 = "insert into 版本_项目表 (规范项目名称,速度目标值,单双线,定额体系,材料信息价,地区,省份,项目类型,人工单价,设计阶段) values ('" + 规范项目名称 + "','" + 速度目标值 + "','" + 单双线 + "','" + 定额体系 + "','" + 材料信息价 + "','" + 地区 + "','" + 省份 + "','" + 项目类型 + "','" + 人工单价 + "','" + 设计阶段 + "')";
                dao2.Execute(sql4);

            }
            if (k >= 0)
            {
                MessageBox.Show("添加成功", "提示");
                return true;
            }
            return false;
        }
    }
}
