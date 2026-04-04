using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class StateService : ParentService

    {
        public List<tbl_state> GetAll(string cnd = "")
        {
            string query = "select * from tbl_state where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_state> dataList = ConvertDataTable<tbl_state>(GetData(query));
            return dataList;
        }
        public List<tbl_state> GetAautoComplete(string cnd = "")
        {
            string query = "select *,state_name as label,state_name as val from tbl_state where 1=1 ";
            query += cnd;
            List<tbl_state> dataList = ConvertDataTable<tbl_state>(GetData(query));
            return dataList;
        }
        public tbl_state? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_state where id=@id", conn);
            cm.Parameters.AddWithValue("@id", id);
            var dataList = ConvertDataTable<tbl_state>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_state data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_state(state_name,cmp_id) Values(@state_name,@cmp_id)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@state_name", data.state_name);
            cm.Parameters.AddWithValue("@cmp_id", 0);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_state data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_state set state_name=@state_name,cmp_id=@cmp_id where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@state_name", data.state_name);
            cm.Parameters.AddWithValue("@cmp_id", 0);

            Execute(cm);
        }
        public void Delete(tbl_state data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_state where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }
    }
}
