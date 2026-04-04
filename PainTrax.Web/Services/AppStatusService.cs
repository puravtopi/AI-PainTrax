using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using Microsoft.AspNetCore.Mvc.Rendering;
using PainTrax.Web.Models;


namespace MS.Services;
public class AppStatusService: ParentService {
	public List<tbl_app_status> GetAll(string cnd="") {
        string query = "select * from tbl_app_status where 1=1";
        if (!string.IsNullOrEmpty(query))
            query += cnd;
        List<tbl_app_status> dataList = ConvertDataTable<tbl_app_status>(GetData(query));
        return dataList;
	}

	public List<SelectListItem> GetAllDropDown(bool select=true, int selected=0)
	{
		var statusData = GetAll();
		var statuslist = new List<SelectListItem>();
		if (select)
		{
			statuslist.Add(new SelectListItem
			{
				Text = "--Select Status--",
				Value = "0"
			});
		}

		foreach (var item in statusData)
		{
			statuslist.Add(new SelectListItem
			{
				Selected = (item.status_id==selected) ,
				Text = item.status,
				Value = item.status_id.ToString()
			});
		}
		return statuslist;
	}

	public tbl_app_status? GetOne(tbl_app_status data) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_app_status where status_id=@id " , conn);
		cm.Parameters.AddWithValue("@id", data.status_id);
		var datalist = ConvertDataTable<tbl_app_status>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}
    public tbl_app_status? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_app_status where status_id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_app_status>(GetData(cm)).FirstOrDefault();
        return datalist;
    }
    public void Insert(tbl_app_status data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_app_status
		(status,cmp_id)Values
				(@status,@cmp_id)",conn);
		cm.Parameters.AddWithValue("@status", data.status);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        Execute(cm);
}
	public void Update(tbl_app_status data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_app_status SET
		status=@status		where status_id=@status_id",conn);
		cm.Parameters.AddWithValue("@status_id", data.status_id);
		cm.Parameters.AddWithValue("@status", data.status);
	Execute(cm);
}
	public void Delete(tbl_app_status data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_app_status
		where status_id=@status_id",conn);
		cm.Parameters.AddWithValue("@status_id", data.status_id);
	Execute(cm);
}

}