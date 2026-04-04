using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class InsuranceStatusService : ParentService
    {
        public List<tbl_insurance_status_type> GetAll(string cnd = "")
        {
            string query = "select * from tbl_insurance_status_type where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_insurance_status_type> dataList = ConvertDataTable<tbl_insurance_status_type>(GetData(query));
            return dataList;
        }

        public tbl_insurance_status_type? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_insurance_status_type where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_insurance_status_type>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_insurance_status_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_insurance_status_type(Name) Values(@Name)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Name", data.Name);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_insurance_status_type data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_insurance_status_type set Name=@Name where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            cm.Parameters.AddWithValue("@Name", data.Name);
            Execute(cm);
        }
       
    }
}
