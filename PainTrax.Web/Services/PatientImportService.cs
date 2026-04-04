using MS.Models;
using MySql.Data.MySqlClient;
using PainTrax.Services;
using System;
using System.Data;

namespace PainTrax.Web.Services
{
    public class PatientImportService : ParentService
    {
        #region Page1
        public int InsertPage1(tbl_ie_page1 data)
        {

            var obj = this.GetOne(data.ie_id.Value);

            if (obj == null)
            {
                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_page1
		(history,pmh,psh,allergies,medication,family_history,social_history,cc,pe,ie_id,patient_id,cmp_id,appt_reason,assessment,occupation,impairment_rating)Values
				(@history,@pmh,@psh,@allergies,@medication,@family_history,@social_history,@cc,@pe,@ie_id,@patient_id,@cmp_id,@appt_reason,@assessment,@occupation,@impairment_rating);select @@identity;", conn);
                cm.Parameters.AddWithValue("@history", data.history);
                cm.Parameters.AddWithValue("@pmh", data.pmh);
                cm.Parameters.AddWithValue("@psh", data.psh);
                cm.Parameters.AddWithValue("@allergies", data.allergies);
                cm.Parameters.AddWithValue("@medication", data.medication);
                cm.Parameters.AddWithValue("@family_history", data.family_history);
                cm.Parameters.AddWithValue("@social_history", data.social_history);
                cm.Parameters.AddWithValue("@cc", data.cc);
                cm.Parameters.AddWithValue("@pe", data.pe);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
                cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
                cm.Parameters.AddWithValue("@assessment", data.assessment);
                cm.Parameters.AddWithValue("@occupation", data.occupation);
                cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);
                var result = ExecuteScalar(cm);
                return result;
            }
            else
            {
                MySqlCommand cm = new MySqlCommand(@"Update tbl_ie_page1 set
		history=@history,pmh=@pmh,psh=@psh,allergies=@allergies,medication=@medication,family_history=@family_history,social_history=@social_history,appt_reason=@appt_reason,
        cc=@cc,pe=@pe,assessment=@assessment,occupation=@occupation,impairment_rating=@impairment_rating where ie_id=@ie_id;select 1;", conn); ;
                cm.Parameters.AddWithValue("@history", data.history);
                cm.Parameters.AddWithValue("@pmh", data.pmh);
                cm.Parameters.AddWithValue("@psh", data.psh);
                cm.Parameters.AddWithValue("@allergies", data.allergies);
                cm.Parameters.AddWithValue("@medication", data.medication);
                cm.Parameters.AddWithValue("@family_history", data.family_history);
                cm.Parameters.AddWithValue("@social_history", data.social_history);
                cm.Parameters.AddWithValue("@cc", data.cc);
                cm.Parameters.AddWithValue("@pe", data.pe);
                cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
                cm.Parameters.AddWithValue("@assessment", data.assessment);
                cm.Parameters.AddWithValue("@occupation", data.occupation);
                cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);

                var result = ExecuteScalar(cm);
                return result;
            }
        }

        public tbl_patient_ie? GetOne(int data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_page1 where ie_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data);
            var datalist = ConvertDataTable<tbl_patient_ie>(GetData(cm)).FirstOrDefault();
            return datalist;
        }


        public int InsertPage1FU(tbl_fu_page1 data)
        {

            var obj = this.GetOneFU(data.fu_id.Value);

            if (obj == null)
            {
                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_fu_page1
		(history,pmh,psh,allergies,medication,family_history,social_history,cc,pe,ie_id,patient_id,cmp_id,assessment,occupation,appt_reason,impairment_rating)Values
				(@history,@pmh,@psh,@allergies,@medication,@family_history,@social_history,@cc,@pe,@ie_id,@patient_id,@cmp_id,@assessment,@occupation,@appt_reason,@impairment_rating);select @@identity;", conn);
                cm.Parameters.AddWithValue("@history", data.history);
                cm.Parameters.AddWithValue("@pmh", data.pmh);
                cm.Parameters.AddWithValue("@psh", data.psh);
                cm.Parameters.AddWithValue("@allergies", data.allergies);
                cm.Parameters.AddWithValue("@medication", data.medication);
                cm.Parameters.AddWithValue("@family_history", data.family_history);
                cm.Parameters.AddWithValue("@social_history", data.social_history);
                cm.Parameters.AddWithValue("@cc", data.cc);
                cm.Parameters.AddWithValue("@pe", data.pe);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
                cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
                cm.Parameters.AddWithValue("@assessment", data.assessment);
                cm.Parameters.AddWithValue("@occupation", data.occupation);
                cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);
                var result = ExecuteScalar(cm);
                return result;
            }
            else
            {
                MySqlCommand cm = new MySqlCommand(@"Update tbl_fu_page1 set
		history=@history,pmh=@pmh,psh=@psh,allergies=@allergies,medication=@medication,family_history=@family_history,social_history=@social_history,appt_reason=@appt_reason,
impairment_rating=@impairment_rating,appt_reason=@appt_reason,
        cc=@cc,pe=@pe,assessment=@assessment,occupation=@occupation,impairment_rating=@impairment_rating where fu_id=@fu_id;select 1;", conn); ;
                cm.Parameters.AddWithValue("@history", data.history);
                cm.Parameters.AddWithValue("@pmh", data.pmh);
                cm.Parameters.AddWithValue("@psh", data.psh);
                cm.Parameters.AddWithValue("@allergies", data.allergies);
                cm.Parameters.AddWithValue("@medication", data.medication);
                cm.Parameters.AddWithValue("@family_history", data.family_history);
                cm.Parameters.AddWithValue("@social_history", data.social_history);
                cm.Parameters.AddWithValue("@cc", data.cc);
                cm.Parameters.AddWithValue("@pe", data.pe);
                cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
                cm.Parameters.AddWithValue("@fu_id", data.fu_id);
                cm.Parameters.AddWithValue("@assessment", data.assessment);
                cm.Parameters.AddWithValue("@occupation", data.occupation);
                cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);

                var result = ExecuteScalar(cm);
                 return result;
            }
        }

        public tbl_patient_ie? GetOneFU(int data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_fu_page1 where fu_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data);
            var datalist = ConvertDataTable<tbl_patient_ie>(GetData(cm)).FirstOrDefault();
            return datalist;
        }


        #endregion

        #region Page2
        public int InsertPage2(tbl_ie_page2 data)
        {

            var obj = this.GetPage2One(data.ie_id.Value);

            if (obj == null)
            {
                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_page2
		(ros,aod,ie_id,patient_id,cmp_id)Values
				(@ros,@aod,@ie_id,@patient_id,@cmp_id);select @@identity;", conn);
                cm.Parameters.AddWithValue("@ros", data.ros);
                cm.Parameters.AddWithValue("@aod", data.aod);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
                var result = ExecuteScalar(cm);
                return result;
            }
            else
            {
                MySqlCommand cm = new MySqlCommand(@"update tbl_ie_page2 set
		ros=@ros,aod=@aod where ie_id=@ie_id;select 1;", conn);
                cm.Parameters.AddWithValue("@ros", data.ros);
                cm.Parameters.AddWithValue("@aod", data.aod);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
              
                var result = ExecuteScalar(cm);
                return result;
            }
        }

        public tbl_ie_page2? GetPage2One(int data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_page2 where ie_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data);
            var datalist = ConvertDataTable<tbl_ie_page2>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int InsertPage2FU(tbl_fu_page2 data)
        {

            var obj = this.GetPage2OneFU(data.fu_id.Value);

            if (obj == null)
            {
                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_fu_page2
		(ros,aod,ie_id,patient_id,cmp_id,fu_id)Values
				(@ros,@aod,@ie_id,@patient_id,@cmp_id,@fu_id);select @@identity;", conn);
                cm.Parameters.AddWithValue("@ros", data.ros);
                cm.Parameters.AddWithValue("@aod", data.aod);
                cm.Parameters.AddWithValue("@ie_id", data.ie_id);
                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
                cm.Parameters.AddWithValue("@fu_id", data.fu_id);
                var result = ExecuteScalar(cm);
                return result;
            }
            else
            {
                MySqlCommand cm = new MySqlCommand(@"update tbl_fu_page2 set
		ros=@ros,aod=@aod where fu_id=@fu_id;select 1;", conn);
                cm.Parameters.AddWithValue("@ros", data.ros);
                cm.Parameters.AddWithValue("@aod", data.aod);
                cm.Parameters.AddWithValue("@fu_id", data.fu_id);

                var result = ExecuteScalar(cm);
                return result;
            }
        }

        public tbl_ie_page2? GetPage2OneFU(int data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_fu_page2 where fu_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data);
            var datalist = ConvertDataTable<tbl_ie_page2>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        #endregion

        #region NE

        public int InsertNE(tbl_ie_ne data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_ie_ne
		(neurological_exam,sensory,manual_muscle_strength_testing,ie_id,patient_id,other_content,cmp_id)Values
				(@neurological_exam,@sensory,@manual_muscle_strength_testing,@ie_id,@patient_id,@other_content,@cmp_id);select @@identity;", conn);
            cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
            cm.Parameters.AddWithValue("@sensory", data.sensory);
            cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
            cm.Parameters.AddWithValue("@ie_id", data.ie_id);
            cm.Parameters.AddWithValue("@other_content", data.other_content);
            cm.Parameters.AddWithValue("@patient_id", data.patient_id);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = ExecuteScalar(cm);
            return result;
        }


        public void InsertNEFU(tbl_fu_ne data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_fu_ne
		(neurological_exam,sensory,manual_muscle_strength_testing,other_content,ie_id,patient_id,cmp_id,fu_id)Values
				(@neurological_exam,@sensory,@manual_muscle_strength_testing,@other_content,@ie_id,@patient_id,@cmp_id,@fu_id)", conn);
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

        public void UpdateNEFU(tbl_fu_ne data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_ne SET
		neurological_exam=@neurological_exam,
		sensory=@sensory,
other_content=@other_content,
		manual_muscle_strength_testing=@manual_muscle_strength_testing
			where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
            cm.Parameters.AddWithValue("@sensory", data.sensory);
            cm.Parameters.AddWithValue("@other_content", data.other_content);
            cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
           

            Execute(cm);
        }

        public int UpdatePageNE(tbl_ie_ne data)
        {
            MySqlCommand cm = new MySqlCommand(@"update tbl_ie_ne
		 set neurological_exam=@neurological_exam,sensory=@sensory,manual_muscle_strength_testing=@manual_muscle_strength_testing,other_content=@other_content
         where id=@id;select 1;", conn);

            cm.Parameters.AddWithValue("@neurological_exam", data.neurological_exam);
            cm.Parameters.AddWithValue("@sensory", data.sensory);
            cm.Parameters.AddWithValue("@other_content", data.other_content);
            cm.Parameters.AddWithValue("@manual_muscle_strength_testing", data.manual_muscle_strength_testing);
            cm.Parameters.AddWithValue("@id", data.id);

            var result = ExecuteScalar(cm);
            return result;
        }

        public tbl_ie_ne GetOneNE(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_ne where ie_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_ie_ne>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public tbl_fu_ne? GetOneNEFU(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_fu_ne where fu_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_fu_ne>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        #endregion
    }
}
