using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PainTrax.Services.Helper
{
    public class dbHelper
    {
        MySqlConnection con;
        MySqlCommand cmd = new MySqlCommand();

        public dbHelper()
        {
            con = new MySqlConnection("Data Source=localhost;Database=dbpaintrax;User ID=root;Password=123456");
        }

        public int executeQuery(string query)
        {
            int val = 1;
            try
            {
                cmd.CommandText = query;
                cmd.Connection = con;
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                val = -1;
            }

            return val;
        }

        public DataSet getDatase(string query)
        {
            DataSet ds = new DataSet(); 
            try
            {
                cmd.CommandText = query;
                cmd.Connection = con;
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(ds);
            }
            catch (Exception ex)
            {
              
            }

            return ds;
        }
    }
}
