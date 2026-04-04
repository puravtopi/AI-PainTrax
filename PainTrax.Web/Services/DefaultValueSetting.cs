using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using PainTrax.Web.Models;


namespace PainTrax.Web.Services
{
    public class DefaultValueSettingServices : ParentService
    {

        #region Page1
        public int InsertPage1(tbl_ie_page1_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_page1_default
		(history,vital,pmh,psh,allergies,medication,family_history,social_history,cc,pe,rom,note,cmp_id,bodypart,daignosis_delimit,daignosis_desc,assessment,plan,dd,work_status,occupation)Values
				(@history,@vital,@pmh,@psh,@allergies,@medication,@family_history,@social_history,@cc,@pe,@rom,@note,@cmp_id,@bodypart,@daignosis_delimit,@daignosis_desc,@assessment,@plan,@dd,@work_status,@occupation);select @@identity;", conn);
            cm.Parameters.AddWithValue("@history", data.history);
            cm.Parameters.AddWithValue("@vital", data.vital);
            cm.Parameters.AddWithValue("@pmh", data.pmh);
            cm.Parameters.AddWithValue("@psh", data.psh);
            cm.Parameters.AddWithValue("@allergies", data.allergies);
            cm.Parameters.AddWithValue("@medication", data.medication);
            cm.Parameters.AddWithValue("@family_history", data.family_history);
            cm.Parameters.AddWithValue("@social_history", data.social_history);
            cm.Parameters.AddWithValue("@cc", data.cc);
            cm.Parameters.AddWithValue("@pe", data.pe);
            cm.Parameters.AddWithValue("@rom", data.rom);
            cm.Parameters.AddWithValue("@note", data.note);

            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);

            cm.Parameters.AddWithValue("@bodypart", data.bodypart);
            cm.Parameters.AddWithValue("@daignosis_desc", data.daignosis_desc);
            cm.Parameters.AddWithValue("@daignosis_delimit", data.daignosis_delimit);
            cm.Parameters.AddWithValue("@assessment", data.assessment);
            cm.Parameters.AddWithValue("@plan", data.plan);
            cm.Parameters.AddWithValue("@dd", data.dd);
            cm.Parameters.AddWithValue("@work_status", data.work_status);
            cm.Parameters.AddWithValue("@occupation", data.occupation);
            var result = ExecuteScalar(cm);
            return result;
        }

        public int UpdatePage1(tbl_ie_page1_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"update tbl_ie_page1_default set
		history=@history,vital=@vital,pmh=@pmh,psh=@psh,allergies=@allergies,medication=@medication,family_history=@family_history,
        social_history=@social_history,cc=@cc,pe=@pe,rom=@rom,note=@note,
        bodypart=@bodypart,daignosis_desc=@daignosis_desc,daignosis_delimit=@daignosis_delimit,occupation=@occupation,
        assessment=@assessment,plan=@plan,dd=@dd,work_status=@work_status 

        where id=@id
				 ;select 1;", conn);
            cm.Parameters.AddWithValue("@history", data.history);
            cm.Parameters.AddWithValue("@vital", data.vital);
            cm.Parameters.AddWithValue("@pmh", data.pmh);
            cm.Parameters.AddWithValue("@psh", data.psh);
            cm.Parameters.AddWithValue("@allergies", data.allergies);
            cm.Parameters.AddWithValue("@medication", data.medication);
            cm.Parameters.AddWithValue("@family_history", data.family_history);
            cm.Parameters.AddWithValue("@social_history", data.social_history);
            cm.Parameters.AddWithValue("@cc", data.cc);
            cm.Parameters.AddWithValue("@pe", data.pe);
            cm.Parameters.AddWithValue("@rom", data.rom);
            cm.Parameters.AddWithValue("@note", data.note);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@bodypart", data.bodypart);
            cm.Parameters.AddWithValue("@daignosis_desc", data.daignosis_desc);
            cm.Parameters.AddWithValue("@daignosis_delimit", data.daignosis_delimit);
            cm.Parameters.AddWithValue("@assessment", data.assessment);
            cm.Parameters.AddWithValue("@plan", data.plan);
            cm.Parameters.AddWithValue("@dd", data.dd);
            cm.Parameters.AddWithValue("@work_status", data.work_status);
            cm.Parameters.AddWithValue("@occupation", data.occupation);
            var result = ExecuteScalar(cm);
            return result;
        }

        public tbl_ie_page1_default? GetOnePage1(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_page1_default where cmp_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_ie_page1_default>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        #endregion

        #region Page2
        public int InsertPage2(tbl_ie_page2_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_page2_default
		(ros,aod,other,cmp_id)Values
				(@ros,@aod,@other,@cmp_id);select @@identity;", conn);

            cm.Parameters.AddWithValue("@ros", data.ros);
            cm.Parameters.AddWithValue("@aod", data.aod);
            cm.Parameters.AddWithValue("@other", data.other);

            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = ExecuteScalar(cm);
            return result;
        }

        public int UpdatePage2(tbl_ie_page2_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"update tbl_ie_page2_default
		 set ros=@ros,aod=@aod,other=@other where id=@id;select 1;", conn);

            cm.Parameters.AddWithValue("@ros", data.ros);
            cm.Parameters.AddWithValue("@aod", data.aod);
            cm.Parameters.AddWithValue("@other", data.other);
            cm.Parameters.AddWithValue("@id", data.id);

            var result = ExecuteScalar(cm);
            return result;
        }
        public tbl_ie_page2_default? GetOnePage2(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_page2_default where cmp_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_ie_page2_default>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        #endregion

        #region NE
        public int InsertNE(tbl_ie_ne_default data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_ne_default
        (neurological_exam,sensory,manual_muscle_strength_testing,other_content,cmp_id)values
        (@neurological_exam,@sensory,@manual_muscle_strength_testing,@other_content,@cmp_id);select 1;", conn);
            
            cm.Parameters.AddWithValue("@neurological_exam",data.neurological_exam);
            cm.Parameters.AddWithValue("@sensory",data.sensory);
            cm.Parameters.AddWithValue("@manual_muscle_strength_testing",data.manual_muscle_strength_testing);
            cm.Parameters.AddWithValue("@other_content", data.other_content);
           
            cm.Parameters.AddWithValue("@cmp_id",data.cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }

        public int UpdateNE(tbl_ie_ne_default data) 
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_ie_ne_default set 
        neurological_exam=@neurological_exam,
        sensory=@sensory,
        manual_muscle_strength_testing=@manual_muscle_strength_testing,
        other_content=@other_content,
       
        cmp_id=@cmp_id where id=@id;select 1;", conn);

            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
            cm.Parameters.AddWithValue("@sensory", data.sensory);
            cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
            cm.Parameters.AddWithValue("@other_content", data.other_content);
           
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }

        public tbl_ie_ne_default? GetOneNE(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_ne_default where cmp_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_ie_ne_default>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        #endregion
    }
}
