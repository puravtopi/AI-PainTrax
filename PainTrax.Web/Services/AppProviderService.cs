using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace MS.Services;
public class AppProviderService: ParentService {
	public List<tbl_app_provider> GetAll(string cnd = "") {
        string query = "select * from tbl_app_provider where 1=1";
        if (!string.IsNullOrEmpty(query))
            query += cnd;
        List<tbl_app_provider> dataList = ConvertDataTable<tbl_app_provider>(GetData(query));
        return dataList;
    }
	public List<SelectListItem> GetAllCompany(int cmdid)
	{
		MySqlCommand cm = new MySqlCommand("select * from tbl_app_provider where cmp_id=@cmp_id ", conn);
		cm.Parameters.AddWithValue("@cmp_id", cmdid);
		var datalist = ConvertDataTable<tbl_app_provider>(GetData(cm));
		var list = new List<SelectListItem>();

		foreach (var item in datalist)
		{
			list.Add(new SelectListItem
			{
				Text = item.provider_name,
				Value = item.provider_id.ToString()
			});
		}

		return list;
		
	}

	public tbl_app_provider? GetOne(tbl_app_provider data) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_app_provider where provider_id=@id " , conn);
		cm.Parameters.AddWithValue("@provider_id", data.provider_id);
		var datalist = ConvertDataTable<tbl_app_provider>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

    public tbl_app_provider? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_app_provider where provider_id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_app_provider>(GetData(cm)).FirstOrDefault();
        return datalist;
    }
    public void Insert(tbl_app_provider data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_app_provider
		(provider_name,cmp_id)Values
				(@provider_name,@cmp_id)",conn);
		cm.Parameters.AddWithValue("@provider_name", data.provider_name);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
		Execute(cm);
}
	public void Update(tbl_app_provider data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_app_provider SET
		provider_name=@provider_name where provider_id=@provider_id", conn);
		cm.Parameters.AddWithValue("@provider_id", data.provider_id);
		cm.Parameters.AddWithValue("@provider_name", data.provider_name);
	//	cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
		Execute(cm);
}
	public void Delete(tbl_app_provider data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_app_provider
		where provider_id=@provider_id",conn);
		cm.Parameters.AddWithValue("@provider_id", data.provider_id);
	Execute(cm);
}

}