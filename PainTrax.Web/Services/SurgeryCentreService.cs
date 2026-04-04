using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class SurgeryCentreService : ParentService
    {
        public List<tbl_surgerycenter> GetAll(string cnd = "")
        {
            string query = "select * from tbl_surgerycenter where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_surgerycenter> dataList = ConvertDataTable<tbl_surgerycenter>(GetData(query));
            return dataList;
        }

        public tbl_surgerycenter? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_surgerycenter where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_surgerycenter>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_surgerycenter data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_surgerycenter(Surgerycenter_name,Address,ContactNo,ContactPerson,cmp_id,is_active) Values(@Surgerycenter_name,@Address,@ContactNo,@ContactPerson,@cmp_id,@is_active)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Surgerycenter_name", data.Surgerycenter_name);
            cm.Parameters.AddWithValue("@Address", data.Address);
            cm.Parameters.AddWithValue("@ContactNo", data.ContactNo);
            cm.Parameters.AddWithValue("@ContactPerson", data.ContactPerson);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@is_active", data.is_active);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_surgerycenter data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_surgerycenter set Surgerycenter_name=@Surgerycenter_name,Address=@Address,ContactNo=@ContactNo,ContactPerson=@ContactPerson,cmp_id=@cmp_id,is_active=@is_active where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Surgerycenter_name", data.Surgerycenter_name);
            cm.Parameters.AddWithValue("@Address", data.Address);
            cm.Parameters.AddWithValue("@ContactNo", data.ContactNo);
            cm.Parameters.AddWithValue("@ContactPerson", data.ContactPerson);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@is_active", data.is_active);
            Execute(cm);
        }
        public void Delete(tbl_surgerycenter data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_surgerycenter set is_active=@is_active where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);            
            cm.Parameters.AddWithValue("@is_active", false);
            Execute(cm);
        }
    }
}
