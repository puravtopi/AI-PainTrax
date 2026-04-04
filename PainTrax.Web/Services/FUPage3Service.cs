using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUPage3Service: ParentService {
	public List<tbl_fu_page3> GetAll() {
		List<tbl_fu_page3> dataList = ConvertDataTable<tbl_fu_page3>(GetData("select * from tbl_fu_page3")); 
	return dataList;
	}

	public tbl_fu_page3? GetOne(int id) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_fu_page3 where fu_id=@id " , conn);
		cm.Parameters.AddWithValue("@id", id);
		var datalist = ConvertDataTable<tbl_fu_page3>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_fu_page3 data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_fu_page3
		(ie_id,cmp_id,patient_id,gait,goal,care,universal,diagcervialbulge_date,diagcervialbulge_study,diagcervialbulge,diagcervialbulge_comma,
		diagcervialbulge_text,diagcervialbulge_hnp1,diagcervialbulge_hnp2,diagthoracicbulge_date,diagthoracicbulge_study,diagthoracicbulge,diagthoracicbulge_comma,
		diagthoracicbulge_text,diagthoracicbulge_hnp1,diagthoracicbulge_hnp2,diaglumberbulge_date,diaglumberbulge_study,diaglumberbulge,diaglumberbulge_comma,
		diaglumberbulge_text,diaglumberbulge_hnp1,diaglumberbulge_hnp2,diagleftshoulder_date,diagleftshoulder_study,diagleftshoulder,diagleftshoulder_comma,
		diagleftshoulder_text,diagrightshoulder_date,diagrightshoulder_study,diagrightshoulder,diagrightshoulder_comma,
		diagrightshoulder_text,diagleftknee_date,diagleftknee_study,diagleftknee,diagleftknee_comma,
		diagleftknee_text,diagrightknee_date,diagrightknee_study,diagrightknee,diagrightknee_comma,
		diagrightknee_text,diagbrain_date,diagbrain_study,diagbrain,diagbrain_comma,
		diagbrain_text,other1_date,other1_study,other1,other1_comma,
		other1_text,other2_date,other2_study,other2,other2_comma,
		other2_text,other3_date,other3_study,other3,other3_comma,
		other3_text,other4_date,other4_study,other4,other4_comma,
		other4_text,other5_date,other5_study,other5,other5_comma,
		other5_text,other6_date,other6_study,other6,other6_comma,
		other6_text,other7_date,other7_study,other7,other7_comma,
		other7_text,followupin,followupin_date,discharge_medications,fu_id)Values
		(@ie_id,@cmp_id,@patient_id,@gait,@goal,@care,@universal,@diagcervialbulge_date,@diagcervialbulge_study,@diagcervialbulge,@diagcervialbulge_comma,
		@diagcervialbulge_text,@diagcervialbulge_hnp1,@diagcervialbulge_hnp2,@diagthoracicbulge_date,@diagthoracicbulge_study,@diagthoracicbulge,@diagthoracicbulge_comma,
		@diagthoracicbulge_text,@diagthoracicbulge_hnp1,@diagthoracicbulge_hnp2,@diaglumberbulge_date,@diaglumberbulge_study,@diaglumberbulge,@diaglumberbulge_comma,
		@diaglumberbulge_text,@diaglumberbulge_hnp1,@diaglumberbulge_hnp2,@diagleftshoulder_date,@diagleftshoulder_study,@diagleftshoulder,@diagleftshoulder_comma,
		@diagleftshoulder_text,@diagrightshoulder_date,@diagrightshoulder_study,@diagrightshoulder,@diagrightshoulder_comma,
		@diagrightshoulder_text,@diagleftknee_date,@diagleftknee_study,@diagleftknee,@diagleftknee_comma,
		@diagleftknee_text,@diagrightknee_date,@diagrightknee_study,@diagrightknee,@diagrightknee_comma,
		@diagrightknee_text,@diagbrain_date,@diagbrain_study,@diagbrain,@diagbrain_comma,
		@diagbrain_text,@other1_date,@other1_study,@other1,@other1_comma,
		@other1_text,@other2_date,@other2_study,@other2,@other2_comma,
		@other2_text,@other3_date,@other3_study,@other3,@other3_comma,
		@other3_text,@other4_date,@other4_study,@other4,@other4_comma,
		@other4_text,@other5_date,@other5_study,@other5,@other5_comma,
		@other5_text,@other6_date,@other6_study,@other6,@other6_comma,
		@other6_text,@other7_date,@other7_study,@other7,@other7_comma,
		@other7_text,@followupin,@followupin_date,@discharge_medications,@fu_id)", conn);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@gait", data.gait);
		cm.Parameters.AddWithValue("@goal", data.goal);
		cm.Parameters.AddWithValue("@care", data.care);
		cm.Parameters.AddWithValue("@universal", data.universal);
		cm.Parameters.AddWithValue("@diagcervialbulge_date", data.diagcervialbulge_date);
		cm.Parameters.AddWithValue("@diagcervialbulge_study", data.diagcervialbulge_study);
		cm.Parameters.AddWithValue("@diagcervialbulge", data.diagcervialbulge);
		cm.Parameters.AddWithValue("@diagcervialbulge_comma", data.diagcervialbulge_comma);
		cm.Parameters.AddWithValue("@diagcervialbulge_text", data.diagcervialbulge_text);
		cm.Parameters.AddWithValue("@diagcervialbulge_hnp1", data.diagcervialbulge_hnp1);
		cm.Parameters.AddWithValue("@diagcervialbulge_hnp2", data.diagcervialbulge_hnp2);
		cm.Parameters.AddWithValue("@diagthoracicbulge_date", data.diagthoracicbulge_date);
		cm.Parameters.AddWithValue("@diagthoracicbulge_study", data.diagthoracicbulge_study);
		cm.Parameters.AddWithValue("@diagthoracicbulge", data.diagthoracicbulge);
		cm.Parameters.AddWithValue("@diagthoracicbulge_comma", data.diagthoracicbulge_comma);
		cm.Parameters.AddWithValue("@diagthoracicbulge_text", data.diagthoracicbulge_text);
		cm.Parameters.AddWithValue("@diagthoracicbulge_hnp1", data.diagthoracicbulge_hnp1);
		cm.Parameters.AddWithValue("@diagthoracicbulge_hnp2", data.diagthoracicbulge_hnp2);
		cm.Parameters.AddWithValue("@diaglumberbulge_date", data.diaglumberbulge_date);
		cm.Parameters.AddWithValue("@diaglumberbulge_study", data.diaglumberbulge_study);
		cm.Parameters.AddWithValue("@diaglumberbulge", data.diaglumberbulge);
		cm.Parameters.AddWithValue("@diaglumberbulge_comma", data.diaglumberbulge_comma);
		cm.Parameters.AddWithValue("@diaglumberbulge_text", data.diaglumberbulge_text);
		cm.Parameters.AddWithValue("@diaglumberbulge_hnp1", data.diaglumberbulge_hnp1);
		cm.Parameters.AddWithValue("@diaglumberbulge_hnp2", data.diaglumberbulge_hnp2);
		cm.Parameters.AddWithValue("@diagleftshoulder_date", data.diagleftshoulder_date);
		cm.Parameters.AddWithValue("@diagleftshoulder_study", data.diagleftshoulder_study);
		cm.Parameters.AddWithValue("@diagleftshoulder", data.diagleftshoulder);
		cm.Parameters.AddWithValue("@diagleftshoulder_comma", data.diagleftshoulder_comma);
		cm.Parameters.AddWithValue("@diagleftshoulder_text", data.diagleftshoulder_text);
		cm.Parameters.AddWithValue("@diagrightshoulder_date", data.diagrightshoulder_date);
		cm.Parameters.AddWithValue("@diagrightshoulder_study", data.diagrightshoulder_study);
		cm.Parameters.AddWithValue("@diagrightshoulder", data.diagrightshoulder);
		cm.Parameters.AddWithValue("@diagrightshoulder_comma", data.diagrightshoulder_comma);
		cm.Parameters.AddWithValue("@diagrightshoulder_text", data.diagrightshoulder_text);
		cm.Parameters.AddWithValue("@diagleftknee_date", data.diagleftknee_date);
		cm.Parameters.AddWithValue("@diagleftknee_study", data.diagleftknee_study);
		cm.Parameters.AddWithValue("@diagleftknee", data.diagleftknee);
		cm.Parameters.AddWithValue("@diagleftknee_comma", data.diagleftknee_comma);
		cm.Parameters.AddWithValue("@diagleftknee_text", data.diagleftknee_text);
		cm.Parameters.AddWithValue("@diagrightknee_date", data.diagrightknee_date);
		cm.Parameters.AddWithValue("@diagrightknee_study", data.diagrightknee_study);
		cm.Parameters.AddWithValue("@diagrightknee", data.diagrightknee);
		cm.Parameters.AddWithValue("@diagrightknee_comma", data.diagrightknee_comma);
		cm.Parameters.AddWithValue("@diagrightknee_text", data.diagrightknee_text);
		cm.Parameters.AddWithValue("@diagbrain_date", data.diagbrain_date);
		cm.Parameters.AddWithValue("@diagbrain_study", data.diagbrain_study);
		cm.Parameters.AddWithValue("@diagbrain", data.diagbrain);
		cm.Parameters.AddWithValue("@diagbrain_comma", data.diagbrain_comma);
		cm.Parameters.AddWithValue("@diagbrain_text", data.diagbrain_text);
		cm.Parameters.AddWithValue("@other1_date", data.other1_date);
		cm.Parameters.AddWithValue("@other1_study", data.other1_study);
		cm.Parameters.AddWithValue("@other1", data.other1);
		cm.Parameters.AddWithValue("@other1_comma", data.other1_comma);
		cm.Parameters.AddWithValue("@other1_text", data.other1_text);
		cm.Parameters.AddWithValue("@other2_date", data.other2_date);
		cm.Parameters.AddWithValue("@other2_study", data.other2_study);
		cm.Parameters.AddWithValue("@other2", data.other2);
		cm.Parameters.AddWithValue("@other2_comma", data.other2_comma);
		cm.Parameters.AddWithValue("@other2_text", data.other2_text);
		cm.Parameters.AddWithValue("@other3_date", data.other3_date);
		cm.Parameters.AddWithValue("@other3_study", data.other3_study);
		cm.Parameters.AddWithValue("@other3", data.other3);
		cm.Parameters.AddWithValue("@other3_comma", data.other3_comma);
		cm.Parameters.AddWithValue("@other3_text", data.other3_text);
		cm.Parameters.AddWithValue("@other4_date", data.other4_date);
		cm.Parameters.AddWithValue("@other4_study", data.other4_study);
		cm.Parameters.AddWithValue("@other4", data.other4);
		cm.Parameters.AddWithValue("@other4_comma", data.other4_comma);
		cm.Parameters.AddWithValue("@other4_text", data.other4_text);
		cm.Parameters.AddWithValue("@other5_date", data.other5_date);
		cm.Parameters.AddWithValue("@other5_study", data.other5_study);
		cm.Parameters.AddWithValue("@other5", data.other5);
		cm.Parameters.AddWithValue("@other5_comma", data.other5_comma);
		cm.Parameters.AddWithValue("@other5_text", data.other5_text);
		cm.Parameters.AddWithValue("@other6_date", data.other6_date);
		cm.Parameters.AddWithValue("@other6_study", data.other6_study);
		cm.Parameters.AddWithValue("@other6", data.other6);
		cm.Parameters.AddWithValue("@other6_comma", data.other6_comma);
		cm.Parameters.AddWithValue("@other6_text", data.other6_text);
		cm.Parameters.AddWithValue("@other7_date", data.other7_date);
		cm.Parameters.AddWithValue("@other7_study", data.other7_study);
		cm.Parameters.AddWithValue("@other7", data.other7);
		cm.Parameters.AddWithValue("@other7_comma", data.other7_comma);
		cm.Parameters.AddWithValue("@other7_text", data.other7_text);
		cm.Parameters.AddWithValue("@followupin", data.followupin);
		cm.Parameters.AddWithValue("@followupin_date", data.followupin_date);
		cm.Parameters.AddWithValue("@discharge_medications", data.discharge_medications);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        Execute(cm);
}
	public void Update(tbl_fu_page3 data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_page3 SET
		ie_id=@ie_id,
		cmp_id=@cmp_id,
		patient_id=@patient_id,
		gait=@gait,
		goal=@goal,
		care=@care,
		universal=@universal,
		diagcervialbulge_date=@diagcervialbulge_date,
		diagcervialbulge_study=@diagcervialbulge_study,
		diagcervialbulge=@diagcervialbulge,
		diagcervialbulge_comma=@diagcervialbulge_comma,
		diagcervialbulge_text=@diagcervialbulge_text,
		diagcervialbulge_hnp1=@diagcervialbulge_hnp1,
		diagcervialbulge_hnp2=@diagcervialbulge_hnp2,
		diagthoracicbulge_date=@diagthoracicbulge_date,
		diagthoracicbulge_study=@diagthoracicbulge_study,
		diagthoracicbulge=@diagthoracicbulge,
		diagthoracicbulge_comma=@diagthoracicbulge_comma,
		diagthoracicbulge_text=@diagthoracicbulge_text,
		diagthoracicbulge_hnp1=@diagthoracicbulge_hnp1,
		diagthoracicbulge_hnp2=@diagthoracicbulge_hnp2,
		diaglumberbulge_date=@diaglumberbulge_date,
		diaglumberbulge_study=@diaglumberbulge_study,
		diaglumberbulge=@diaglumberbulge,
		diaglumberbulge_comma=@diaglumberbulge_comma,
		diaglumberbulge_text=@diaglumberbulge_text,
		diaglumberbulge_hnp1=@diaglumberbulge_hnp1,
		diaglumberbulge_hnp2=@diaglumberbulge_hnp2,
		diagleftshoulder_date=@diagleftshoulder_date,
		diagleftshoulder_study=@diagleftshoulder_study,
		diagleftshoulder=@diagleftshoulder,
		diagleftshoulder_comma=@diagleftshoulder_comma,
		diagleftshoulder_text=@diagleftshoulder_text,
		diagrightshoulder_date=@diagrightshoulder_date,
		diagrightshoulder_study=@diagrightshoulder_study,
		diagrightshoulder=@diagrightshoulder,
		diagrightshoulder_comma=@diagrightshoulder_comma,
		diagrightshoulder_text=@diagrightshoulder_text,
		diagleftknee_date=@diagleftknee_date,
		diagleftknee_study=@diagleftknee_study,
		diagleftknee=@diagleftknee,
		diagleftknee_comma=@diagleftknee_comma,
		diagleftknee_text=@diagleftknee_text,
		diagrightknee_date=@diagrightknee_date,
		diagrightknee_study=@diagrightknee_study,
		diagrightknee=@diagrightknee,
		diagrightknee_comma=@diagrightknee_comma,
		diagrightknee_text=@diagrightknee_text,
		diagbrain_date=@diagbrain_date,
		diagbrain_study=@diagbrain_study,
		diagbrain=@diagbrain,
		diagbrain_comma=@diagbrain_comma,
		diagbrain_text=@diagbrain_text,
		other1_date=@other1_date,
		other1_study=@other1_study,
		other1=@other1,
		other1_comma=@other1_comma,
		other1_text=@other1_text,
		other2_date=@other2_date,
		other2_study=@other2_study,
		other2=@other2,
		other2_comma=@other2_comma,
		other2_text=@other2_text,
		other3_date=@other3_date,
		other3_study=@other3_study,
		other3=@other3,
		other3_comma=@other3_comma,
		other3_text=@other3_text,
		other4_date=@other4_date,
		other4_study=@other4_study,
		other4=@other4,
		other4_comma=@other4_comma,
		other4_text=@other4_text,
		other5_date=@other5_date,
		other5_study=@other5_study,
		other5=@other5,
		other5_comma=@other5_comma,
		other5_text=@other5_text,
		other6_date=@other6_date,
		other6_study=@other6_study,
		other6=@other6,
		other6_comma=@other6_comma,
		other6_text=@other6_text,
		other7_date=@other7_date,
		other7_study=@other7_study,
		other7=@other7,
		other7_comma=@other7_comma,
		other7_text=@other7_text,
		followupin=@followupin,
		followupin_date=@followupin_date,
		discharge_medications=@discharge_medications		where id=@id", conn);
		cm.Parameters.AddWithValue("@id", data.id);
		cm.Parameters.AddWithValue("@ie_id", data.ie_id);
		cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@gait", data.gait);
		cm.Parameters.AddWithValue("@goal", data.goal);
		cm.Parameters.AddWithValue("@care", data.care);
		cm.Parameters.AddWithValue("@universal", data.universal);
		cm.Parameters.AddWithValue("@diagcervialbulge_date", data.diagcervialbulge_date);
		cm.Parameters.AddWithValue("@diagcervialbulge_study", data.diagcervialbulge_study);
		cm.Parameters.AddWithValue("@diagcervialbulge", data.diagcervialbulge);
		cm.Parameters.AddWithValue("@diagcervialbulge_comma", data.diagcervialbulge_comma);
		cm.Parameters.AddWithValue("@diagcervialbulge_text", data.diagcervialbulge_text);
		cm.Parameters.AddWithValue("@diagcervialbulge_hnp1", data.diagcervialbulge_hnp1);
		cm.Parameters.AddWithValue("@diagcervialbulge_hnp2", data.diagcervialbulge_hnp2);
		cm.Parameters.AddWithValue("@diagthoracicbulge_date", data.diagthoracicbulge_date);
		cm.Parameters.AddWithValue("@diagthoracicbulge_study", data.diagthoracicbulge_study);
		cm.Parameters.AddWithValue("@diagthoracicbulge", data.diagthoracicbulge);
		cm.Parameters.AddWithValue("@diagthoracicbulge_comma", data.diagthoracicbulge_comma);
		cm.Parameters.AddWithValue("@diagthoracicbulge_text", data.diagthoracicbulge_text);
		cm.Parameters.AddWithValue("@diagthoracicbulge_hnp1", data.diagthoracicbulge_hnp1);
		cm.Parameters.AddWithValue("@diagthoracicbulge_hnp2", data.diagthoracicbulge_hnp2);
		cm.Parameters.AddWithValue("@diaglumberbulge_date", data.diaglumberbulge_date);
		cm.Parameters.AddWithValue("@diaglumberbulge_study", data.diaglumberbulge_study);
		cm.Parameters.AddWithValue("@diaglumberbulge", data.diaglumberbulge);
		cm.Parameters.AddWithValue("@diaglumberbulge_comma", data.diaglumberbulge_comma);
		cm.Parameters.AddWithValue("@diaglumberbulge_text", data.diaglumberbulge_text);
		cm.Parameters.AddWithValue("@diaglumberbulge_hnp1", data.diaglumberbulge_hnp1);
		cm.Parameters.AddWithValue("@diaglumberbulge_hnp2", data.diaglumberbulge_hnp2);
		cm.Parameters.AddWithValue("@diagleftshoulder_date", data.diagleftshoulder_date);
		cm.Parameters.AddWithValue("@diagleftshoulder_study", data.diagleftshoulder_study);
		cm.Parameters.AddWithValue("@diagleftshoulder", data.diagleftshoulder);
		cm.Parameters.AddWithValue("@diagleftshoulder_comma", data.diagleftshoulder_comma);
		cm.Parameters.AddWithValue("@diagleftshoulder_text", data.diagleftshoulder_text);
		cm.Parameters.AddWithValue("@diagrightshoulder_date", data.diagrightshoulder_date);
		cm.Parameters.AddWithValue("@diagrightshoulder_study", data.diagrightshoulder_study);
		cm.Parameters.AddWithValue("@diagrightshoulder", data.diagrightshoulder);
		cm.Parameters.AddWithValue("@diagrightshoulder_comma", data.diagrightshoulder_comma);
		cm.Parameters.AddWithValue("@diagrightshoulder_text", data.diagrightshoulder_text);
		cm.Parameters.AddWithValue("@diagleftknee_date", data.diagleftknee_date);
		cm.Parameters.AddWithValue("@diagleftknee_study", data.diagleftknee_study);
		cm.Parameters.AddWithValue("@diagleftknee", data.diagleftknee);
		cm.Parameters.AddWithValue("@diagleftknee_comma", data.diagleftknee_comma);
		cm.Parameters.AddWithValue("@diagleftknee_text", data.diagleftknee_text);
		cm.Parameters.AddWithValue("@diagrightknee_date", data.diagrightknee_date);
		cm.Parameters.AddWithValue("@diagrightknee_study", data.diagrightknee_study);
		cm.Parameters.AddWithValue("@diagrightknee", data.diagrightknee);
		cm.Parameters.AddWithValue("@diagrightknee_comma", data.diagrightknee_comma);
		cm.Parameters.AddWithValue("@diagrightknee_text", data.diagrightknee_text);
		cm.Parameters.AddWithValue("@diagbrain_date", data.diagbrain_date);
		cm.Parameters.AddWithValue("@diagbrain_study", data.diagbrain_study);
		cm.Parameters.AddWithValue("@diagbrain", data.diagbrain);
		cm.Parameters.AddWithValue("@diagbrain_comma", data.diagbrain_comma);
		cm.Parameters.AddWithValue("@diagbrain_text", data.diagbrain_text);
		cm.Parameters.AddWithValue("@other1_date", data.other1_date);
		cm.Parameters.AddWithValue("@other1_study", data.other1_study);
		cm.Parameters.AddWithValue("@other1", data.other1);
		cm.Parameters.AddWithValue("@other1_comma", data.other1_comma);
		cm.Parameters.AddWithValue("@other1_text", data.other1_text);
		cm.Parameters.AddWithValue("@other2_date", data.other2_date);
		cm.Parameters.AddWithValue("@other2_study", data.other2_study);
		cm.Parameters.AddWithValue("@other2", data.other2);
		cm.Parameters.AddWithValue("@other2_comma", data.other2_comma);
		cm.Parameters.AddWithValue("@other2_text", data.other2_text);
		cm.Parameters.AddWithValue("@other3_date", data.other3_date);
		cm.Parameters.AddWithValue("@other3_study", data.other3_study);
		cm.Parameters.AddWithValue("@other3", data.other3);
		cm.Parameters.AddWithValue("@other3_comma", data.other3_comma);
		cm.Parameters.AddWithValue("@other3_text", data.other3_text);
		cm.Parameters.AddWithValue("@other4_date", data.other4_date);
		cm.Parameters.AddWithValue("@other4_study", data.other4_study);
		cm.Parameters.AddWithValue("@other4", data.other4);
		cm.Parameters.AddWithValue("@other4_comma", data.other4_comma);
		cm.Parameters.AddWithValue("@other4_text", data.other4_text);
		cm.Parameters.AddWithValue("@other5_date", data.other5_date);
		cm.Parameters.AddWithValue("@other5_study", data.other5_study);
		cm.Parameters.AddWithValue("@other5", data.other5);
		cm.Parameters.AddWithValue("@other5_comma", data.other5_comma);
		cm.Parameters.AddWithValue("@other5_text", data.other5_text);
		cm.Parameters.AddWithValue("@other6_date", data.other6_date);
		cm.Parameters.AddWithValue("@other6_study", data.other6_study);
		cm.Parameters.AddWithValue("@other6", data.other6);
		cm.Parameters.AddWithValue("@other6_comma", data.other6_comma);
		cm.Parameters.AddWithValue("@other6_text", data.other6_text);
		cm.Parameters.AddWithValue("@other7_date", data.other7_date);
		cm.Parameters.AddWithValue("@other7_study", data.other7_study);
		cm.Parameters.AddWithValue("@other7", data.other7);
		cm.Parameters.AddWithValue("@other7_comma", data.other7_comma);
		cm.Parameters.AddWithValue("@other7_text", data.other7_text);
		cm.Parameters.AddWithValue("@followupin", data.followupin);
		cm.Parameters.AddWithValue("@followupin_date", data.followupin_date);
		cm.Parameters.AddWithValue("@discharge_medications", data.discharge_medications);
	Execute(cm);
}
	public void Delete(tbl_fu_page3 data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_page3
		where id=@id",conn);
		cm.Parameters.AddWithValue("@id", data.id);
	Execute(cm);
}

}