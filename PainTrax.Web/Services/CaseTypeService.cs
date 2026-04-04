using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class CaseTypeService : ParentService
    {
        public List<tbl_casetype> GetAll(string cnd = "")
        {
            string query = "select * from tbl_casetype where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_casetype> dataList = ConvertDataTable<tbl_casetype>(GetData(query));
            return dataList;
        }

        public tbl_casetype? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_casetype where Id=@id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_casetype>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_casetype data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_casetype(casetype,cmp_id) Values(@casetype,@cmp_id)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@casetype", data.casetype);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_casetype data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_casetype set casetype=@casetype,cmp_id=@cmp_id where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@casetype", data.casetype);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void Delete(tbl_casetype data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_casetype where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }
    }
}
