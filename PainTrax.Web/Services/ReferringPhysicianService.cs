using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class ReferringPhysicianService : ParentService
    {
        public List<tbl_referring_physician> GetAll(string cnd = "")
        {
            string query = "select * from tbl_referring_physician where 1=1";
            if (!string.IsNullOrEmpty(query))
                query += cnd;
            List<tbl_referring_physician> dataList = ConvertDataTable<tbl_referring_physician>(GetData(query));
            return dataList;
        }

        public tbl_referring_physician? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_referring_physician where Id=@id", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var dataList = ConvertDataTable<tbl_referring_physician>(GetData(cm)).FirstOrDefault();
            return dataList;
        }

        public int Insert(tbl_referring_physician data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_referring_physician(physicianname,address,locationid,cmp_id)
                Values(@physicianname,@address,@locationid,@cmp_id)", conn);
            //cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@physicianname", data.physicianname);
            cm.Parameters.AddWithValue("@address",data.address);
            cm.Parameters.AddWithValue("@locationid", data.locationid);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = Execute(cm);
            return result;
        }
        public void Update(tbl_referring_physician data)
        {
            MySqlCommand cm = new MySqlCommand(@"Update tbl_referring_physician set 
                physicianname=@physicianname,
                address=@address,
                locationid=@locationid,
                cmp_id=@cmp_id 

                where id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@physicianname", data.physicianname);
            cm.Parameters.AddWithValue("@address", data.address);
            cm.Parameters.AddWithValue("@locationid", data.locationid);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void Delete(tbl_referring_physician data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_referring_physician where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }
    }
}
