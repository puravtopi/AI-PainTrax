using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class AppRangeService: ParentService {
	public List<tbl_app_range> GetAll() {
		List<tbl_app_range> dataList = ConvertDataTable<tbl_app_range>(GetData("select * from tbl_app_range")); 
	return dataList;
	}

	public tbl_app_range? GetOne(tbl_app_range data) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_app_range where id=@id " , conn);
		cm.Parameters.AddWithValue("@id", data.id);
		var datalist = ConvertDataTable<tbl_app_range>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_app_range data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_app_range
		(range_start,rnage_end)Values
				(@range_start,@rnage_end)",conn);
		cm.Parameters.AddWithValue("@range_start", data.range_start);
		cm.Parameters.AddWithValue("@rnage_end", data.rnage_end);
	Execute(cm);
}
	public void Update(tbl_app_range data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_app_range SET
		range_start=@range_start,
		rnage_end=@rnage_end		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@range_start", data.range_start);
		cm.Parameters.AddWithValue("@rnage_end", data.rnage_end);
	Execute(cm);
}
	public void Delete(tbl_app_range data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_app_range
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}