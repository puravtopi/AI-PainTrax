using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class AttorneysService : ParentService
    {

        public AttorneysService()
        {
        }

        public List<tbl_attorneys> GetAll(string cnd = "")
        {
            string query = "select * from tbl_attorneys where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            query = query+" order by attorney";

            List<tbl_attorneys> dataList = ConvertDataTable<tbl_attorneys>(GetData(query));
            return dataList;
        }

        public List<tbl_attorneys> GetAautoComplete(string cnd = "")
        {
            string query = "select *,attorney as label,attorney as val from tbl_attorneys where 1=1 ";
            query += cnd;
            List<tbl_attorneys> dataList = ConvertDataTable<tbl_attorneys>(GetData(query));
            return dataList;
        }


        public tbl_attorneys? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_attorneys where Id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<tbl_attorneys>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int Insert(tbl_attorneys data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_attorneys
		(Attorney,Address,City,State,Zip,ContactName,ContactNo,EmailId,CreatedDate,CreatedBy,old_id,cmp_id,Paralegal)Values
				(@Attorney,@Address,@City,@State,@Zip,@ContactName,@ContactNo,@EmailId,@CreatedDate,@CreatedBy,@old_id,@cmp_id,@Paralegal);select @@identity;", conn);
            cm.Parameters.AddWithValue("@Attorney", data.Attorney);
            cm.Parameters.AddWithValue("@Address", data.Address);
            cm.Parameters.AddWithValue("@City", data.City);
            cm.Parameters.AddWithValue("@State", data.State);
            cm.Parameters.AddWithValue("@Zip", data.Zip);
            cm.Parameters.AddWithValue("@ContactName", data.ContactName);
            cm.Parameters.AddWithValue("@ContactNo", data.ContactNo);
            cm.Parameters.AddWithValue("@EmailId", data.EmailId);
            cm.Parameters.AddWithValue("@CreatedDate", System.DateTime.Now);
            cm.Parameters.AddWithValue("@CreatedBy", data.CreatedBy);
            cm.Parameters.AddWithValue("@old_id", data.old_id);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@Paralegal", data.Paralegal);
            var result = ExecuteScalar(cm);
            return result;
        }
        public void Update(tbl_attorneys data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_attorneys SET
		Attorney=@Attorney,
		Address=@Address,
		City=@City,
		State=@State,
		Zip=@Zip,
		ContactName=@ContactName,
		ContactNo=@ContactNo,
		EmailId=@EmailId,
		CreatedDate=@CreatedDate,
		CreatedBy=@CreatedBy,
Paralegal=@Paralegal,
		old_id=@old_id,
		cmp_id=@cmp_id		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Attorney", data.Attorney);
            cm.Parameters.AddWithValue("@Address", data.Address);
            cm.Parameters.AddWithValue("@City", data.City);
            cm.Parameters.AddWithValue("@State", data.State);
            cm.Parameters.AddWithValue("@Zip", data.Zip);
            cm.Parameters.AddWithValue("@ContactName", data.ContactName);
            cm.Parameters.AddWithValue("@ContactNo", data.ContactNo);
            cm.Parameters.AddWithValue("@EmailId", data.EmailId);
            cm.Parameters.AddWithValue("@CreatedDate", data.CreatedDate);
            cm.Parameters.AddWithValue("@CreatedBy", data.CreatedBy);
            cm.Parameters.AddWithValue("@old_id", data.old_id);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@Paralegal", data.Paralegal);
            Execute(cm);
        }
        public void Delete(tbl_attorneys data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_attorneys
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }

    }
}