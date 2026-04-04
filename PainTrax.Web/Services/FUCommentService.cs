using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUCommentService: ParentService {
	public List<tbl_fu_comment> GetAll() {
		List<tbl_fu_comment> dataList = ConvertDataTable<tbl_fu_comment>(GetData("select * from tbl_fu_comment")); 
	return dataList;
	}

	public tbl_fu_comment? GetOne(int id) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_fu_comment where fu_id=@id " , conn);
		cm.Parameters.AddWithValue("@id", id);
		var datalist = ConvertDataTable<tbl_fu_comment>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_fu_comment data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_fu_comment
		(ie_id,patient_id,comment_details,fu_id)Values
				(@ie_id,@patient_id,@comment_details,@fu_id)", conn);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@comment_details", data.comment_details);
		cm.Parameters.AddWithValue("@fu_id", data.fu_id);
	Execute(cm);
}
	public void Update(tbl_fu_comment data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_comment SET
		ie_id=@ie_id,
	
		patient_id=@patient_id,
		comment_details=@comment_details		where id=@id", conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@comment_details", data.comment_details);
       
        Execute(cm);
}
	public void Delete(tbl_fu_comment data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_comment
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}