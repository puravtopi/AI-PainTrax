using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class LocationsService : ParentService
    {
        public List<tbl_locations> GetAll(string cnd = "")
        {
            string query = "select * from tbl_locations where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;


            List<tbl_locations> dataList = ConvertDataTable<tbl_locations>(GetData(query));
            return dataList;
        }

        public tbl_locations? GetOne(tbl_locations data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_locations where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_locations>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public void Insert(tbl_locations data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_locations
		(location,setasdefault,address,city,state,zipcode,emailid,telephone,contactpersonname,nameofpractice,fax,drfname,drlname,isactive,cmp_id,createddate,createdby,header_template)Values
				(@location,@setasdefault,@address,@city,@state,@zipcode,@emailid,@telephone,@contactpersonname,@nameofpractice,@fax,@drfname,@drlname,@isactive,@cmp_id,@createddate,@createdby,@header_template);select @@identity;", conn);
            cm.Parameters.AddWithValue("@location", data.location);
            cm.Parameters.AddWithValue("@setasdefault", data.setasdefault);
            cm.Parameters.AddWithValue("@address", data.address);
            cm.Parameters.AddWithValue("@city", data.city);
            cm.Parameters.AddWithValue("@state", data.state);
            cm.Parameters.AddWithValue("@zipcode", data.zipcode);
            cm.Parameters.AddWithValue("@emailid", data.emailid);
            cm.Parameters.AddWithValue("@telephone", data.telephone);
            cm.Parameters.AddWithValue("@contactpersonname", data.contactpersonname);
            cm.Parameters.AddWithValue("@nameofpractice", data.nameofpractice);
            cm.Parameters.AddWithValue("@fax", data.fax);
            cm.Parameters.AddWithValue("@drfname", data.drfname);
            cm.Parameters.AddWithValue("@drlname", data.drlname);
            cm.Parameters.AddWithValue("@isactive", data.isactive);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@createddate", data.createddate);
            cm.Parameters.AddWithValue("@createdby", data.createdby);
            cm.Parameters.AddWithValue("@header_template", data.header_template);

            int id = ExecuteScalar(cm);

            addlocationintogroup(data.cmp_id.Value, data.location, id);
        }
        public void Update(tbl_locations data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_locations SET
		location=@location,
		setasdefault=@setasdefault,
		address=@address,
		city=@city,
		state=@state,
		zipcode=@zipcode,
		emailid=@emailid,
		telephone=@telephone,
		contactpersonname=@contactpersonname,
		nameofpractice=@nameofpractice,
		fax=@fax,
		drfname=@drfname,
		drlname=@drlname,
		isactive=@isactive,
		
		updatedate=@updatedate,
		updatedby=@updatedby,
		header_template=@header_template
	 	where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@location", data.location);
            cm.Parameters.AddWithValue("@setasdefault", data.setasdefault);
            cm.Parameters.AddWithValue("@address", data.address);
            cm.Parameters.AddWithValue("@city", data.city);
            cm.Parameters.AddWithValue("@state", data.state);
            cm.Parameters.AddWithValue("@zipcode", data.zipcode);
            cm.Parameters.AddWithValue("@emailid", data.emailid);
            cm.Parameters.AddWithValue("@telephone", data.telephone);
            cm.Parameters.AddWithValue("@contactpersonname", data.contactpersonname);
            cm.Parameters.AddWithValue("@nameofpractice", data.nameofpractice);
            cm.Parameters.AddWithValue("@fax", data.fax);
            cm.Parameters.AddWithValue("@drfname", data.drfname);
            cm.Parameters.AddWithValue("@drlname", data.drlname);
            cm.Parameters.AddWithValue("@isactive", data.isactive);
            cm.Parameters.AddWithValue("@updatedate", data.updatedate);
            cm.Parameters.AddWithValue("@updatedby", data.updatedby);
            cm.Parameters.AddWithValue("@header_template", data.header_template);
            Execute(cm);
        }
        public void Delete(tbl_locations data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_locations
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }

        private void addlocationintogroup(int cmpid, string loc_name, int loca_id)
        {
            string query = "select * from tbl_groups where cmp_id=" + cmpid;


            List<tbl_groups> dataList = ConvertDataTable<tbl_groups>(GetData(query));

            foreach (var item in dataList)
            {
                var location_name = item.location_name + "," + loc_name;
                var location_ids = item.Location_ids + "," + loca_id;

                query = "update tbl_groups set location_name='" + location_name + "',location_ids='" + location_ids + "' where Id=" + item.Id;

                MySqlCommand cm = new MySqlCommand(query, conn);
                Execute(cm);

            }

        }

    }
}