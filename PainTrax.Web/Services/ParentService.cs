using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PainTrax.Services
{
    public class ParentService
    {
        protected MySqlConnection conn = new MySqlConnection();

        public ParentService()
        {
            IConfiguration myConfig = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appSettings.json")
            .Build();

            var config = myConfig.GetSection("ConnectionStrings");

            //conn = new MySqlConnection("Data Source=10.10.93.18;port=2109;Database=dbpaintrax_live;User ID=purav;Password=G0d$peed@123");
            //Live
            //conn = new MySqlConnection("Data Source=10.10.93.20;port=3306;Database=dbpaintrax_live;User ID=purav;Password=amd@2006");
            //Local
            // conn = new MySqlConnection("Data Source=localhost;Database=dbpaintrax_live;User ID=root;Password=");
             conn = new MySqlConnection(config["MySQLConString"]);
           
        }

        protected static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        protected static List<T> SPConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        protected static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {

                    // if (pro.PropertyType == typeof(Boolean))
                    if (pro.PropertyType.ToString().Contains("System.Boolean"))
                    {
                        if (pro.Name.ToLower() == column.ColumnName.ToLower())
                            pro.SetValue(obj, dr[column.ColumnName] == DBNull.Value ? null : (dr[column.ColumnName].ToString() == "1" ? true : false), null);
                        else
                            continue;

                    }
                    else
                    {
                        if (pro.Name.ToLower() == column.ColumnName.ToLower())
                            pro.SetValue(obj, dr[column.ColumnName] == DBNull.Value ? null : dr[column.ColumnName], null);
                        else
                            continue;
                    }
                }
            }
            return obj;
        }

        public int Execute(MySqlCommand cmd)
        {
            try
            {
                conn.Open();
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return 0;
        }

        public int ExecuteScalar(MySqlCommand cmd)
        {

            try
            {
                conn.Open();
                int r = 0;
                object result = cmd.ExecuteScalar();
                if (result != null)
                {
                    r = Convert.ToInt32(result);
                }
                return r;
            }
            catch (Exception ex)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return 0;
        }

        public DataTable GetData(string query)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand(query, conn);
            try
            {
                conn.Open();
                MySqlDataReader datar = cm.ExecuteReader();
                if (datar.HasRows)
                {
                    dt.Load(datar);
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return dt;
        }

        public DataTable GetData(MySqlCommand cm)
        {
            DataTable dt = new DataTable();
            cm.Connection = conn;
            try
            {
                conn.Open();
                MySqlDataReader datar = cm.ExecuteReader();
                if (datar.HasRows)
                {
                    dt.Load(datar);
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return dt;
        }

        public string ExecuteSP(string sp, List<MySqlParameter> param)
        {

            MySqlCommand sqlCommand = new MySqlCommand(sp, conn);
            sqlCommand.CommandType = CommandType.StoredProcedure;
            foreach (MySqlParameter item in param)
            {
                sqlCommand.Parameters.Add(item);
            }

            try
            {
                conn.Open();
                return sqlCommand.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
            finally
            {
                conn.Close();
            }
        }

    }
}
