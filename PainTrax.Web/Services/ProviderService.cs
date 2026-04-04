using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services;
public class ProviderService: ParentService {
	public List<tbl_provider> GetAll(string cnd = "") {
        string query = "select * from tbl_provider where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_provider> dataList = ConvertDataTable<tbl_provider>(GetData(query));
        return dataList;
    }

	public tbl_provider? GetOne(tbl_provider data) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_provider where id=@id " , conn);
		cm.Parameters.AddWithValue("@id", data.id);
		var datalist = ConvertDataTable<tbl_provider>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_provider data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_provider
		(provider,address1,address2,city,state,zip,emailaddress,telephone,contactperson,set_as_default,fax,cmp_id,created_by,created_date)Values
				(@provider,@address1,@address2,@city,@state,@zip,@emailaddress,@telephone,@contactperson,@set_as_default,@fax,@cmp_id,@created_by,@created_date)",conn);
		cm.Parameters.AddWithValue("@provider", data.provider);
		cm.Parameters.AddWithValue("@address1", data.address1);
		cm.Parameters.AddWithValue("@address2", data.address2);
		cm.Parameters.AddWithValue("@city", data.city);
		cm.Parameters.AddWithValue("@state", data.state);
		cm.Parameters.AddWithValue("@zip", data.zip);
		cm.Parameters.AddWithValue("@emailaddress", data.emailaddress);
		cm.Parameters.AddWithValue("@telephone", data.telephone);
		cm.Parameters.AddWithValue("@contactperson", data.contactperson);
		cm.Parameters.AddWithValue("@set_as_default", data.set_as_default);
		cm.Parameters.AddWithValue("@fax", data.fax);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
		cm.Parameters.AddWithValue("@created_by", data.created_by);
	
		cm.Parameters.AddWithValue("@created_date", data.created_date);
		
	Execute(cm);
}
	public void Update(tbl_provider data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_provider SET
		provider=@provider,
		address1=@address1,
		address2=@address2,
		city=@city,
		state=@state,
		zip=@zip,
		emailaddress=@emailaddress,
		telephone=@telephone,
		contactperson=@contactperson,
		set_as_default=@set_as_default,
		fax=@fax,
	
		
		updated_by=@updated_by,
	
		updated_date=@updated_date		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@provider", data.provider);
		cm.Parameters.AddWithValue("@address1", data.address1);
		cm.Parameters.AddWithValue("@address2", data.address2);
		cm.Parameters.AddWithValue("@city", data.city);
		cm.Parameters.AddWithValue("@state", data.state);
		cm.Parameters.AddWithValue("@zip", data.zip);
		cm.Parameters.AddWithValue("@emailaddress", data.emailaddress);
		cm.Parameters.AddWithValue("@telephone", data.telephone);
		cm.Parameters.AddWithValue("@contactperson", data.contactperson);
		cm.Parameters.AddWithValue("@set_as_default", data.set_as_default);
		cm.Parameters.AddWithValue("@fax", data.fax);
		
	
		cm.Parameters.AddWithValue("@updated_by", data.updated_by);
	
		cm.Parameters.AddWithValue("@updated_date", data.updated_date);
	Execute(cm);
}
	public void Delete(tbl_provider data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_provider
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}