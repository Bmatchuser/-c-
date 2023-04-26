using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;
namespace 串口助手
{
    class Dao
    {
        
        public SqlConnection connect()
        {
            string str = @"Data Source=DESKTOP-6J4D6RC;Initial Catalog=DB_test;Integrated Security=True";
            SqlConnection sc = new SqlConnection(str);
            sc.Open(); //打开数据库连接
            return sc;
        }
        public SqlCommand command(string sql)
        {
            SqlCommand cmd = new SqlCommand(sql, connect());
            return cmd;
        }
        //用于 update delete insert，返回受影响的行数
        public int Execute(string sql)
        {
            return command(sql).ExecuteNonQuery();
        }
        //用于select，返回sqldateReader对象，包含select到的数据
        public SqlDataReader read(string sql)
        {
            return command(sql).ExecuteReader();
        }
    }
}
