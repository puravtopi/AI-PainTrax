using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using MS.Services;
using PainTrax.Services;

namespace MS.Services;
public class FUNEService: ParentService
{
	public List<tbl_fu_ne> GetAll() {
		List<tbl_fu_ne> dataList = ConvertDataTable<tbl_fu_ne>(GetData("select * from tbl_fu_ne")); 
	return dataList;
	}

	public tbl_fu_ne? GetOne(int id) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_fu_ne where fu_id=@id " , conn);
		cm.Parameters.AddWithValue("@id", id);
		var datalist = ConvertDataTable<tbl_fu_ne>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_fu_ne data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_fu_ne
		(neurological_exam,sensory,manual_muscle_strength_testing,other_content,ie_id,patient_id,cmp_id,fu_id)Values
				(@neurological_exam,@sensory,@manual_muscle_strength_testing,@other_content,@ie_id,@patient_id,@cmp_id,@fu_id)",conn);
		cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
		cm.Parameters.AddWithValue("@sensory", data.sensory);
		cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
		cm.Parameters.AddWithValue("@other_content", data.other_content);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        Execute(cm);
}
	public void Update(tbl_fu_ne data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_ne SET
		neurological_exam=@neurological_exam,
		sensory=@sensory,
		manual_muscle_strength_testing=@manual_muscle_strength_testing,
		other_content=@other_content,
		ie_id=@ie_id,

		patient_id=@patient_id,
		cmp_id=@cmp_id		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
		cm.Parameters.AddWithValue("@sensory", data.sensory);
		cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
		cm.Parameters.AddWithValue("@other_content", data.other_content);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);

        Execute(cm);
}
	public void Delete(tbl_fu_ne data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_ne
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}