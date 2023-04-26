using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
namespace 串口助手
{
    class Dao3
    {
        public SqlConnection connect(string str)
        {
            SqlConnection sc = new SqlConnection(str);
            sc.Open(); //打开数据库连接
            return sc;
        }
        public SqlCommand command(string sql, string str)
        {
            SqlCommand cmd = new SqlCommand(sql, connect(str));
            return cmd;
        }
        //用于 update delete insert，返回受影响的行数
        public int Execute(string sql, string str)
        {
            return command(sql, str).ExecuteNonQuery();
        }
        //用于select，返回sqldateReader对象，包含select到的数据
        public SqlDataReader read(string sql, string str)
        {
            return command(sql, str).ExecuteReader();
        }
        public DataTable GetTable(string sql, string str)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter adapter = null;
            adapter = new SqlDataAdapter(sql, connect(str));
            adapter.Fill(ds);
            return ds.Tables[0];

        }
    }
}
