using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUPage2Service : ParentService {
	public List<tbl_fu_page2> GetAll() {
		List<tbl_fu_page2> dataList = ConvertDataTable<tbl_fu_page2>(GetData("select * from tbl_fu_page2")); 
	return dataList;
	}

	public tbl_fu_page2? GetOne(int id) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_fu_page2 where fu_id=@id " , conn);
		cm.Parameters.AddWithValue("@id", id);
		var datalist = ConvertDataTable<tbl_fu_page2>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_fu_page2 data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_fu_page2
		(ros,aod,other,ie_id,patient_id,cmp_id,fu_id)Values
				(@ros,@aod,@other,@ie_id,@patient_id,@cmp_id,@fu_id)",conn);
		cm.Parameters.AddWithValue("@ros", data.ros);
		cm.Parameters.AddWithValue("@aod", data.aod);
		cm.Parameters.AddWithValue("@other", data.other);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        Execute(cm);
}
	public void Update(tbl_fu_page2 data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_page2 SET
		ros=@ros,
		aod=@aod,
		other=@other,
		ie_id=@ie_id,
		patient_id=@patient_id,
		cmp_id=@cmp_id		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@ros", data.ros);
		cm.Parameters.AddWithValue("@aod", data.aod);
		cm.Parameters.AddWithValue("@other", data.other);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
	Execute(cm);
}
	public void Delete(tbl_fu_page2 data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_page2
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}