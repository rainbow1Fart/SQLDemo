using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace ToolDemo
{
    public class MSSQLHelper
    {
        public MSSQLHelper()
        {
        }

        ~MSSQLHelper()
        {
            if (conn == null)
            {
                return;
            }
        }

        //public string connString = "Data Source=服务器名;Initial Catalog=数据库名;User ID=用户名;Pwd=密码";

        //创建连接对象的变量  
        public SqlConnection conn;

        public bool Connetction(string connString)
        {
            try
            {
                conn = new SqlConnection(connString);
                if (conn == null)
                {
                    return false;
                }

                conn.Open();
                Console.WriteLine(conn.State);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }

        // 执行对数据表中数据的增加、删除、修改操作  
        public int NonQuery(string sql)
        {
            int a = -1;
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                a = cmd.ExecuteNonQuery();
                conn.Close();      //关闭数据库 
            }
            catch (Exception ex)
            {
                conn.Close();      //关闭数据库 
                MessageBox.Show(ex.Message);
                return -1;
            }
            return a;

        }
        // 执行对数据表中数据的查询操作  
        public DataSet Query(string sql)
        {
            DataSet ds = new DataSet();
            try
            {
                if(conn.State == ConnectionState.Closed)
                    conn.Open();
                SqlDataAdapter adp = new SqlDataAdapter(sql, conn);
                adp.Fill(ds);
                conn.Close();      //关闭数据库 
            }
            catch (Exception ex)
            {
                conn.Close();      //关闭数据库 
                MessageBox.Show(ex.Message);
                return null;
            }
            return ds;
        }

    }
}
