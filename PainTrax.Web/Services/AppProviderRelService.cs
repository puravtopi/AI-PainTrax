using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class AppProviderRelService: ParentService {
	public List<tbl_appprovider_rel> GetAll() {
		List<tbl_appprovider_rel> dataList = ConvertDataTable<tbl_appprovider_rel>(GetData("select * from tbl_appprovider_rel")); 
	return dataList;
	}
	public List<tbl_appprovider_rel> GetProvidersByAppoitment(int appointment_id)
	{
		List<tbl_appprovider_rel> dataList = ConvertDataTable<tbl_appprovider_rel>(GetData("select * from tbl_appprovider_rel"));
		return dataList;
	}

	public tbl_appprovider_rel? GetOne(tbl_appprovider_rel data) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_appprovider_rel where id=@id " , conn);
		cm.Parameters.AddWithValue("@id", data.id);
		var datalist = ConvertDataTable<tbl_appprovider_rel>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_appprovider_rel data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_appprovider_rel
		(app_id,app_provider_id)Values
				(@app_id,@app_provider_id)",conn);
		cm.Parameters.AddWithValue("@app_id", data.app_id);
		cm.Parameters.AddWithValue("@app_provider_id", data.app_provider_id);
	Execute(cm);
}
	public void Update(tbl_appprovider_rel data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_appprovider_rel SET
		app_id=@app_id,
		app_provider_id=@app_provider_id		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@app_id", data.app_id);
		cm.Parameters.AddWithValue("@app_provider_id", data.app_provider_id);
	Execute(cm);
}
	public void Delete(tbl_appprovider_rel data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_appprovider_rel
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		Execute(cm);
	}
	public void DeleteByAppointmentId(int appid)
	{
		MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_appprovider_rel
		where app_id=@app_id", conn);
		cm.Parameters.AddWithValue("@app_id", appid);
		Execute(cm);
	}

}