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
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            User user = new User();
            user.Show();
            user.Width = 1200;
            user.Height = 700;
            int height = System.Windows.Forms.SystemInformation.WorkingArea.Height;
            int width = System.Windows.Forms.SystemInformation.WorkingArea.Width;
            int formheight = user.Size.Height;
            int formwidth = user.Size.Width;
            int newformx = width / 2 - formwidth / 2;
            int newformy = height / 2 - formheight / 2;
            user.SetDesktopLocation(newformx, newformy);
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //admin_login admin_login = new admin_login();
            //admin_login.Show();
            //this.Hide();
        }

        private void login_Load(object sender, EventArgs e)
        {

        }
    }
}
