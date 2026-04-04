using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class GroupsService : ParentService
    {
        public List<tbl_groups> GetAll(string cnd = "")
        {

            string query = "select * from tbl_groups where 1=1 ";

            query += cnd;
            List<tbl_groups> dataList = ConvertDataTable<tbl_groups>(GetData(query));
            return dataList;
        }
        public tbl_groups? GetOne(tbl_groups data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_groups where Id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            var datalist = ConvertDataTable<tbl_groups>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        public void Insert(tbl_groups data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_groups
		(Title,Location_ids,CreatedDate,CreatedBy,cmp_id,location_name,pages_ids,pages_name,reports_ids,report_name,role_ids,role_name,form_name)Values
				(@Title,@Location_ids,@CreatedDate,@CreatedBy,@cmp_id,@location_name,@pages_ids,@pages_name,@reports_ids,@report_name,@role_ids,@role_name,@form_name)", conn);
            cm.Parameters.AddWithValue("@Title", data.Title);
            cm.Parameters.AddWithValue("@Location_ids", data.Location_ids);
            cm.Parameters.AddWithValue("@CreatedDate", data.CreatedDate);
            cm.Parameters.AddWithValue("@CreatedBy", data.CreatedBy);
            cm.Parameters.AddWithValue("@location_name", data.location_name);
            cm.Parameters.AddWithValue("@pages_ids", data.pages_ids);
            cm.Parameters.AddWithValue("@pages_name", data.pages_name);
            cm.Parameters.AddWithValue("@reports_ids", data.reports_ids);
            cm.Parameters.AddWithValue("@report_name", data.report_name);
            cm.Parameters.AddWithValue("@role_ids", data.role_ids);
            cm.Parameters.AddWithValue("@role_name", data.role_name);
            cm.Parameters.AddWithValue("@form_name", data.form_name);

            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void Update(tbl_groups data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_groups SET
		Title=@Title,
		Location_ids=@Location_ids,
        location_name=@location_name,
        pages_ids=@pages_ids,
		pages_name=@pages_name,
		reports_ids=@reports_ids,
		report_name=@report_name,
		role_ids=@role_ids,
		role_name=@role_name,
		form_name=@form_name
			where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@Title", data.Title);
            cm.Parameters.AddWithValue("@Location_ids", data.Location_ids);
            cm.Parameters.AddWithValue("@location_name", data.location_name);
            cm.Parameters.AddWithValue("@pages_ids", data.pages_ids);
            cm.Parameters.AddWithValue("@pages_name", data.pages_name);
            cm.Parameters.AddWithValue("@reports_ids", data.reports_ids);
            cm.Parameters.AddWithValue("@report_name", data.report_name);
            cm.Parameters.AddWithValue("@role_ids", data.role_ids);
            cm.Parameters.AddWithValue("@role_name", data.role_name);
            cm.Parameters.AddWithValue("@form_name", data.form_name);

            Execute(cm);
        }
        public void Delete(tbl_groups data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_groups
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }

    }
}