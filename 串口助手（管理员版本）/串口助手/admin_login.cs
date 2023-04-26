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
    public partial class admin_login : Form
    {
        public admin_login()
        {
            InitializeComponent();
        }

        //定时器事件，图片移动到中间的位置以后自动跳转窗口
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (pictureBox1.Location.X < 150)
            {
                pictureBox1.Location = new Point(pictureBox1.Location.X + 1, pictureBox1.Location.Y);
            }
            else
            {
                Admin1 admin1 = new Admin1();
                admin1.Show();//显示这个窗体

                admin1.Width = 1200;
                admin1.Height = 700;
                int height = System.Windows.Forms.SystemInformation.WorkingArea.Height;
                int width = System.Windows.Forms.SystemInformation.WorkingArea.Width;
                int formheight = admin1.Size.Height;
                int formwidth = admin1.Size.Width;
                int newformx = width / 2 - formwidth / 2;
                int newformy = height / 2 - formheight / 2;
                admin1.SetDesktopLocation(newformx, newformy);

                this.Hide();//隐藏现在的窗体
                //this.Close();//关闭现在的窗体
                timer1.Stop();
            }
        }
        //如果通过了登录验证，就将所有的控件隐藏，同时触发timer事件
        private void button1_Click(object sender, EventArgs e)
        {
            if (login())
           {
               timer1.Start();//启动计时器控件，图片开始移动
               textBox2.Visible = false;
               label2.Visible = false;
               button1.Visible = false;;
               button2.Visible = false;
               button3.Visible = false;
           }
        }
        //登录验证
        private bool login()
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("密码不能为空！","提示",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return false;
            }
            
            string sql = "select * from 管理员 where 密码='"+textBox2.Text+"'";
            Dao2 dao2 = new Dao2();
            IDataReader dr = dao2.read(sql);
            if (dr.Read())
            {
                return true;
            }
            else
            {
                MessageBox.Show("密码错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }
        //取消事件
        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = null;
            
        }

        private void admin_login_Load(object sender, EventArgs e)
        {

        }

        private void admin_login_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();//结束整个程序
        }

        private void admin_login_FormClosed_1(object sender, FormClosedEventArgs e)
        {
            Application.Exit();//结束整个程序
        }
        //返回事件
        private void button3_Click(object sender, EventArgs e)
        {
            login login = new login();
            login.Show();
            this.Hide();
        }
        // 回车键登录
        private void admin_login_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                button1_Click(sender, e);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
