using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;

using PainTrax.Web.Models;
using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Services
{
    public class InscosService : ParentService
    {
        public List<tbl_inscos> GetAll(string cnd = "")
        {
            string query = "select * from tbl_inscos where 1=1 ";
            query += cnd;
            List<tbl_inscos> dataList = ConvertDataTable<tbl_inscos>(GetData(query));
            return dataList;
        }

        public List<tbl_inscos> GetAautoComplete(string cnd = "")
        {
            string query = "select distinct cmpname as label,id as ival from tbl_inscos where isactive=1 ";
            query += cnd + " order by cmpname";
           
            List<tbl_inscos> dataList = ConvertDataTable<tbl_inscos>(GetData(query));
            return dataList;
        }

        public tbl_inscos? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_inscos where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_inscos>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int Insert(tbl_inscos data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_inscos
		(cmpname,address1,address2,city,state,emailid,telephone,contactpersonname,setasdefault,isactive,cmp_id,createddate,createdby,zipcode,faxno)Values
				(@cmpname,@address1,@address2,@city,@state,@emailid,@telephone,@contactpersonname,@setasdefault,@isactive,@cmp_id,@createddate,@createdby,@zipcode,@faxno);select @@identity", conn);
            cm.Parameters.AddWithValue("@cmpname", data.cmpname);
            cm.Parameters.AddWithValue("@address1", data.address1);
            cm.Parameters.AddWithValue("@address2", data.address2);
            cm.Parameters.AddWithValue("@city", data.city);
            cm.Parameters.AddWithValue("@state", data.state);
            cm.Parameters.AddWithValue("@emailid", data.emailid);
            cm.Parameters.AddWithValue("@telephone", data.telephone);
            cm.Parameters.AddWithValue("@contactpersonname", data.contactpersonname);
            cm.Parameters.AddWithValue("@setasdefault", data.setasdefault);
            cm.Parameters.AddWithValue("@isactive", data.isactive);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@zipcode", data.zipcode);
            cm.Parameters.AddWithValue("@createddate", System.DateTime.Now);
            cm.Parameters.AddWithValue("@createdby", data.createdby);
            cm.Parameters.AddWithValue("@faxno", data.faxno);

            var result = ExecuteScalar(cm);
            return result;
        }
        public void Update(tbl_inscos data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_inscos SET
		address1=@address1,
		address2=@address2,
		city=@city,
		state=@state,
		emailid=@emailid,
		telephone=@telephone,
		contactpersonname=@contactpersonname,
		setasdefault=@setasdefault,
		isactive=@isactive,
		faxno=@faxno,
		cmpname=@cmpname,
		updatedate=@updatedate,
		updatedby=@updatedby
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@cmpname", data.cmpname);
            cm.Parameters.AddWithValue("@address1", data.address1);
            cm.Parameters.AddWithValue("@address2", data.address2);
            cm.Parameters.AddWithValue("@city", data.city);
            cm.Parameters.AddWithValue("@state", data.state);
            cm.Parameters.AddWithValue("@emailid", data.emailid);
            cm.Parameters.AddWithValue("@telephone", data.telephone);
            cm.Parameters.AddWithValue("@contactpersonname", data.contactpersonname);
            cm.Parameters.AddWithValue("@setasdefault", data.setasdefault);
            cm.Parameters.AddWithValue("@isactive", data.isactive);
            cm.Parameters.AddWithValue("@faxno", data.faxno);
            cm.Parameters.AddWithValue("@zipcode", data.zipcode);
            cm.Parameters.AddWithValue("@updatedate", data.updatedate);
            cm.Parameters.AddWithValue("@updatedby", data.updatedby);

            Execute(cm);
        }
        public void Delete(tbl_inscos data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_inscos
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }

    }
}