using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using PainTrax.Web.ViewModel;
using System.Data;

namespace PainTrax.Web.Services
{
    public class IEToFUService : ParentService
    {
        public List<tbl_ie_fu_select> GetAll(string cnd = "")
        {
            string query = "select * from tbl_ie_fu_select where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_ie_fu_select> dataList = ConvertDataTable<tbl_ie_fu_select>(GetData(query));
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
        public int Insert(IEToFuVM data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_fu_select(tbl_name,tbl_column,cmp_id,created_by,created_date) Values(@tbl_name,@tbl_column,@cmp_id,@created_by,CURDATE())", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@tbl_name", data.tbl_name);
            cm.Parameters.AddWithValue("@tbl_column", data.tbl_column);
            cm.Parameters.AddWithValue("@created_by", data.created_by);
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
        public void Delete(int Id)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_ie_fu_select where cmp_id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", Id);
            Execute(cm);
        }
    }
}
