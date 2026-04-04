using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class DefaultDataServices : ParentService
    {
        public DefaultDataServices()
        {
        }

        public List<tbl_default> GetAll(string cnd = "")
        {
            string query = "select * from tbl_default where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;

            List<tbl_default> dataList = ConvertDataTable<tbl_default>(GetData(query));
            return dataList;
        }

        public List<tbl_default> GetAautoComplete(string cnd = "")
        {
            string query = "select *,attorney as label,attorney as val from tbl_default where 1=1 ";
            query += cnd;
            List<tbl_default> dataList = ConvertDataTable<tbl_default>(GetData(query));
            return dataList;
        }


        public tbl_default? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_default where Id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<tbl_default>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public tbl_default? GetOneByCompany(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_default where cmp_id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<tbl_default>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        public void Delete(tbl_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_default
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            Execute(cm);
        }

    }
}
