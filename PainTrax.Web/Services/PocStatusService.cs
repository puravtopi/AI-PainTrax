using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class PocStatusService : ParentService
    {
        public List<tbl_poc_status_type> GetAll(string cnd = "")
        {
            string query = "select * from tbl_poc_status_type where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_poc_status_type> dataList = ConvertDataTable<tbl_poc_status_type>(GetData(query));
            return dataList;
        }

        public tbl_poc_status_type? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_poc_status_type where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_poc_status_type>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_poc_status_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_poc_status_type(Name,color) Values(@Name,@color)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Name", data.Name);
            cm.Parameters.AddWithValue("@color", data.color);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_poc_status_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_poc_status_type set Name=@Name , color=@color where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            cm.Parameters.AddWithValue("@Name", data.Name);
            cm.Parameters.AddWithValue("@color", data.color);
            Execute(cm);
        }
       
    }
}
