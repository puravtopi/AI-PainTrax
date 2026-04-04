using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class AccidentTypeService: ParentService
    {
        public List<tbl_accidenttype> GetAll(string cnd = "")
        {
            string query = "select * from tbl_accidenttype where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_accidenttype> dataList = ConvertDataTable<tbl_accidenttype>(GetData(query));
            return dataList;
        }

        public tbl_accidenttype? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_accidenttype where Id=@id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_accidenttype>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_accidenttype data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_accidenttype(accidenttype,cmp_id) Values(@accidenttype,@cmp_id)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@accidenttype", data.accidenttype);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_accidenttype data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_accidenttype set accidenttype=@accidenttype,cmp_id=@cmp_id where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@accidenttype", data.accidenttype);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void Delete(tbl_accidenttype data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_accidenttype where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }
    }
}
