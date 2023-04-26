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
    public partial class Admin1 : Form
    {
        //定义两个全局变量，向admin2传递数据
        private string toadmin21; //传递项目号
        private string toadmin22; //传递建设项目名称
        private string toadmin23; //传b版本
        public Admin1()
        {
            InitializeComponent();
        }

        private void Admin1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();//结束整个程序
        }
       
        Dao2 dao2 = new Dao2();
        Dao1 dao1 = new Dao1();
        //界面打开以后就直接给下拉列表赋值
        private void Admin1_Load(object sender, EventArgs e)
        {
            string sql = "select 版本号,版本ID from 版本表";
            DataTable dt = dao2.GetTable(sql);
            this.comboBox1.DataSource = dt;
            this.comboBox1.ValueMember = "版本ID";  //实际值
            this.comboBox1.DisplayMember = "版本号"; //显示值
        }
        //comboBox1点击以后触发事件，自动在dataGridView1显示相应的数据
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
           //string str = "server=DESKTOP-6J4D6RC;database=RecoData2017;uid=sa;pwd=123456";//SQLserver方式连接数据库
            dataGridView1.Rows.Clear();//清空表里的数据
            //通过版本号作为条件查询到对应的版本ID
            string sql2 = "select * from 版本表 where 版本号 = '" + this.comboBox1.Text + "'";
            IDataReader dr2 = dao2.read(sql2);
            while (dr2.Read())
            {
                string str = dr2["版本ID"].ToString();
                //将IP存入txt文件
                Test test = new Test();
                test.WriteText("版本选择.txt",str);
            }
            dr2.Close();

            string sql = "select * from 项目信息";
            IDataReader dr = dao1.read(sql);
            while (dr.Read())
            {
                string 项目编号, 编制办法文号, 建设项目名称, 简称, 设计阶段, 编制范围, 工程总量, 单位, 项目负责人, 概算总值, 概算指标, 标准定额应用, 火车运输标准, 项目版本号, 创建时间, 材料库, 台班库, 设备库, 审查状态, 编制年至开工年年限, 项目密码, 铁路等级, 正线数目, 牵引种类, 闭塞方式, 项目简介, 速度目标值, 打印编制复核, 项目类型, 单位换算;
                项目编号 = dr["项目编号"].ToString();
                建设项目名称 = dr["建设项目名称"].ToString();
                设计阶段 = dr["设计阶段"].ToString();
                编制范围 = dr["编制范围"].ToString();
                工程总量 = dr["工程总量"].ToString();
                单位 = dr["单位"].ToString();
                项目负责人 = dr["项目负责人"].ToString();

                string[] str = { 项目编号, 建设项目名称, 项目负责人, 设计阶段, 编制范围, 单位,工程总量 };
                dataGridView1.Rows.Add(str);
            }   
            dr.Close();  
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        //选择某版本，某项目，点击后跳转界面
        private void button1_Click(object sender, EventArgs e)
        {
            //通过版本号作为条件查询到对应的版本ID
            string sql2 = "select * from 版本表 where 版本号 = '" + this.comboBox1.Text + "'";
            IDataReader dr2 = dao2.read(sql2);
            while (dr2.Read())
            {
                string str = dr2["版本ID"].ToString();
                //将IP存入txt文件
                Test test = new Test();
                test.WriteText("版本选择.txt", str);

                toadmin21 = dataGridView1.SelectedCells[0].Value.ToString();
                toadmin22 = dataGridView1.SelectedCells[1].Value.ToString();
                toadmin23 = dr2["版本号"].ToString();
                //string sql = "select * from 项目信息 where 项目编号  ";
                //IDataReader dr = dao1.read(sql);
                //while (dr.Read())
                //{
                //    string 建设项目名称;
                //    toadmin22 = dr["建设项目名称"].ToString();   
                //}
                //dr.Close(); 
                Admin2 admin2 = new Admin2(toadmin21,toadmin22,toadmin23);

                admin2.Width = 1200;
                admin2.Height = 700;
                int height = System.Windows.Forms.SystemInformation.WorkingArea.Height;
                int width = System.Windows.Forms.SystemInformation.WorkingArea.Width;
                int formheight = admin2.Size.Height;
                int formwidth = admin2.Size.Width;
                int newformx = width / 2 - formwidth / 2;
                int newformy = height / 2 - formheight / 2;
                admin2.SetDesktopLocation(newformx, newformy);

                admin2.Show();
                
            }
            
            
            this.Hide();
            dr2.Close();     
        }
        //搜索功能，如果textBox1不为空，点击搜索后触发搜索功能
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                select();
            }
            else 
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
        }
        //根据textBox1的内容进行查询
        public void select() 
        {
            dataGridView1.Rows.Clear();
            //获取到用户输入的查询信息
            string textbox1 = this.textBox1.Text.Trim();
            string sql = "select * from 项目信息 where  建设项目名称 like '%" + textbox1 + "%' or 项目负责人 like '%" + textbox1 + "%' ";
            IDataReader dr = dao1.read(sql);
            while (dr.Read())
            {
                string 项目编号, 编制办法文号, 建设项目名称, 简称, 设计阶段, 编制范围, 工程总量, 单位, 项目负责人, 概算总值, 概算指标, 标准定额应用, 火车运输标准, 项目版本号, 创建时间, 材料库, 台班库, 设备库, 审查状态, 编制年至开工年年限, 项目密码, 铁路等级, 正线数目, 牵引种类, 闭塞方式, 项目简介, 速度目标值, 打印编制复核, 项目类型, 单位换算;
                项目编号 = dr["项目编号"].ToString();
                建设项目名称 = dr["建设项目名称"].ToString();
                设计阶段 = dr["设计阶段"].ToString();
                编制范围 = dr["编制范围"].ToString();
                工程总量 = dr["工程总量"].ToString();
                单位 = dr["单位"].ToString();
                项目负责人 = dr["项目负责人"].ToString();

                string[] str = { 项目编号, 建设项目名称, 项目负责人, 设计阶段, 编制范围, 单位, 工程总量 };
                dataGridView1.Rows.Add(str);
            }
            dr.Close();  
        }
        //退出
        private void button3_Click(object sender, EventArgs e)
        {
            admin_login admin_login = new admin_login();
            admin_login.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                select2();
            }
            else
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
        }
        //根据textBox2的内容进行查询
        public void select2()
        {
            dataGridView1.Rows.Clear();
            //获取到用户输入的查询信息
            string textbox2 = this.textBox2.Text.Trim();
            string sql = "select * from 项目信息 where  设计阶段 like '%" + textbox2 + "%' ";
            IDataReader dr = dao1.read(sql);
            while (dr.Read())
            {
                string 项目编号, 编制办法文号, 建设项目名称, 简称, 设计阶段, 编制范围, 工程总量, 单位, 项目负责人, 概算总值, 概算指标, 标准定额应用, 火车运输标准, 项目版本号, 创建时间, 材料库, 台班库, 设备库, 审查状态, 编制年至开工年年限, 项目密码, 铁路等级, 正线数目, 牵引种类, 闭塞方式, 项目简介, 速度目标值, 打印编制复核, 项目类型, 单位换算;
                项目编号 = dr["项目编号"].ToString();
                建设项目名称 = dr["建设项目名称"].ToString();
                设计阶段 = dr["设计阶段"].ToString();
                编制范围 = dr["编制范围"].ToString();
                工程总量 = dr["工程总量"].ToString();
                单位 = dr["单位"].ToString();
                项目负责人 = dr["项目负责人"].ToString();

                string[] str = { 项目编号, 建设项目名称, 项目负责人, 设计阶段, 编制范围, 单位, 工程总量 };
                dataGridView1.Rows.Add(str);
            }
            dr.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }
    }
}
