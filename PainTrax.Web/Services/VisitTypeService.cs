using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class VisitTypeService : ParentService
    {
        public List<tbl_visit_type> GetAll(string cnd = "")
        {
            string query = "select * from tbl_visit_type where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_visit_type> dataList = ConvertDataTable<tbl_visit_type>(GetData(query));
            return dataList;
        }

        public tbl_visit_type? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_visit_type where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_visit_type>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_visit_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_visit_type(type) Values(@type)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@type", data.type);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_visit_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_visit_type set type=@type where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            cm.Parameters.AddWithValue("@type", data.type);
            Execute(cm);
        }
       
    }
}
