using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class CompanyServices : ParentService
    {
        public List<tbl_company> GetAll(string cnd = "")
        {
            string query = "select * from tbl_company where 1=1 ";
            query += cnd;

            List<tbl_company> dataList = ConvertDataTable<tbl_company>(GetData(query));
            return dataList;
        }

        public tbl_company? GetOne(tbl_company data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_company where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_company>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public void Insert(tbl_company data)
        {
            try
            {


                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_company
		        (name,address,phone,email,is_active,client_code,created_date,created_by,cmp_type)Values
				(@name,@address,@phone,@email,@is_active,@client_code,@createddate,@createdby,@cmp_type);select @@identity", conn);
                cm.Parameters.AddWithValue("@name", data.name);
                cm.Parameters.AddWithValue("@address", data.address);
                cm.Parameters.AddWithValue("@phone", data.phone);
                cm.Parameters.AddWithValue("@email", data.email);
                cm.Parameters.AddWithValue("@is_active", data.is_active);
                cm.Parameters.AddWithValue("@client_code", data.client_code);
                cm.Parameters.AddWithValue("@createddate", System.DateTime.Now);
                cm.Parameters.AddWithValue("@createdby", 1);
                cm.Parameters.AddWithValue("@cmp_type", data.cmp_type);

                int result = ExecuteScalar(cm);

                cm = new MySqlCommand(@"INSERT INTO tbl_users
		        (fname,lname,emailid,address,fullname,uname,password,groupid,desigid,cmp_id)Values
				(@fname,@lname,@emailid,@address,@fullname,@uname,@password,@groupid,@desigid,@cmp_id);select @@identity;", conn);
                cm.Parameters.AddWithValue("@fname", "admin");
                cm.Parameters.AddWithValue("@lname", "admin");
                cm.Parameters.AddWithValue("@emailid", data.email);
                cm.Parameters.AddWithValue("@address", data.address);
                cm.Parameters.AddWithValue("@fullname", data.name);
                cm.Parameters.AddWithValue("@uname", data.username);
                cm.Parameters.AddWithValue("@password", data.password);
                cm.Parameters.AddWithValue("@groupid", 11);
                cm.Parameters.AddWithValue("@desigid", 23);
                cm.Parameters.AddWithValue("@cmp_id", result);

                int result1 = ExecuteScalar(cm);

                cm = new MySqlCommand(@"INSERT INTO tbl_settings
		        (page_size,location,dateformat,cmp_id)Values
				(@page_size,@location,@dateformat,@cmp_id);select @@identity;", conn);
                cm.Parameters.AddWithValue("@page_size", 50);
                cm.Parameters.AddWithValue("@location",1);
                cm.Parameters.AddWithValue("@dateformat",  "MM/dd/yyyy");
                cm.Parameters.AddWithValue("@cmp_id", result);


                Execute(cm);


            }
            catch (Exception ex)
            {

            }

        }
        public void Update(tbl_company data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_company SET
		name=@name,
	    address=@address,
		phone=@phone,
		email=@email,
		is_active=@is_active,
		client_code=@client_code,
        cmp_type=@cmp_type
	 	where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@name", data.name);
            cm.Parameters.AddWithValue("@address",data.address);
            cm.Parameters.AddWithValue("@phone", data.phone);
            cm.Parameters.AddWithValue("@email", data.email);
            cm.Parameters.AddWithValue("@is_active", data.is_active);
            cm.Parameters.AddWithValue("@client_code", data.client_code);
            cm.Parameters.AddWithValue("@cmp_type", data.cmp_type);

            Execute(cm);
        }
        public void Delete(tbl_company data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_company
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }
    }
}
