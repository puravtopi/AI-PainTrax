using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services;
public class AadjusterService : ParentService
{
    public List<tbl_adjuster> GetAll(string cnd = "")
    {

        string query = "select * from tbl_adjuster where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_adjuster> dataList = ConvertDataTable<tbl_adjuster>(GetData(query));
        return dataList;
    }

    public List<tbl_adjuster> GetAautoComplete(string cnd = "")
    {

        string query = "select *,adjuster as label,adjuster as val from tbl_adjuster where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_adjuster> dataList = ConvertDataTable<tbl_adjuster>(GetData(query));
        return dataList;
    }

    public tbl_adjuster? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_adjuster where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_adjuster>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public int Insert(tbl_adjuster data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_adjuster
		(adjuster,address1,address2,city,state,zip,emailaddress,telephone,contactperson,set_as_default,fax,cmp_id,created_by,created_date)Values
				(@adjuster,@address1,@address2,@city,@state,@zip,@emailaddress,@telephone,@contactperson,@set_as_default,@fax,@cmp_id,@created_by,@created_date);select @@identity;", conn);
        cm.Parameters.AddWithValue("@adjuster", data.adjuster);
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
        cm.Parameters.AddWithValue("@created_date", System.DateTime.Now);

        var result = ExecuteScalar(cm);
        return result;
    }
    public void Update(tbl_adjuster data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_adjuster SET
		adjuster=@adjuster,
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
	
		updated_date=@updated_date		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@adjuster", data.adjuster);
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
    public void Delete(tbl_adjuster data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_adjuster
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}