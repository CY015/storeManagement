using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace 订货系统
{
    class ConnSQL
    {
        // 个人连接SQL Server信息
        private string MySqlCon = "Data Source=DESKTOP-5NV5EFJ;Initial Catalog=OrderSystem;Integrated Security=True";

        //查询
        public DataTable ExecuteQuery(string sqlStr)
        {
            SqlConnection con = new SqlConnection(@MySqlCon);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sqlStr;
            DataTable dt = new DataTable();
            SqlDataAdapter msda;
            msda = new SqlDataAdapter(cmd);
            msda.Fill(dt);
            //con.Close();

            return dt;
        }

        //增删改
        public int ExecuteUpdate(string sqlStr)
        {
            SqlConnection con = new SqlConnection(@MySqlCon);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sqlStr;
            int iud = 0;
            iud = cmd.ExecuteNonQuery();
            //con.Close();

            return iud;
        }
    }
}
