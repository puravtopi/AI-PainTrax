using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class TranscribeService : ParentService
    {

        public int Insert(tbl_transcribe data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_transcribe
        (content,history,cc,plan,subjective,assessment,ie_id,fu_id,length,intake_id,cmp_id,date,file_path,
        ActivitiesOfDailyLiving,Allergies,CareRestrictions,Causation,CoordinationExamination,Diagnoses,
        DiagnosticStudies,FollowUp,Gait,Goals,Medications,MotorExamination,MusculoskeletalExamination,
        NeurologicalExamination,PastMedicalHistory,PastSurgicalHistory,Occupation,PatientInformation,
        Precautions,ProceduresPerformed,RangeOfMotion,Recommendations,ReflexExamination,
        ReviewOfSystems,SensoryExamination,SocialHistory,VitalSigns,WorkAndAdministrativeDetails)
        VALUES
        (@content,@history,@cc,@plan,@subjective,@assessment,@ie_id,@fu_id,@length,@intake_id,@cmp_id,@date,@file_path,
        @ActivitiesOfDailyLiving,@Allergies,@CareRestrictions,@Causation,@CoordinationExamination,@Diagnoses,
        @DiagnosticStudies,@FollowUp,@Gait,@Goals,@Medications,@MotorExamination,@MusculoskeletalExamination,
        @NeurologicalExamination,@PastMedicalHistory,@PastSurgicalHistory,@Occupation,@PatientInformation,
        @Precautions,@ProceduresPerformed,@RangeOfMotion,@Recommendations,@ReflexExamination,
        @ReviewOfSystems,@SensoryExamination,@SocialHistory,@VitalSigns,@WorkAndAdministrativeDetails);
        SELECT @@IDENTITY;", conn);

            // Initial parameters
            cm.Parameters.AddWithValue("@content", (object)data.content ?? DBNull.Value);
            cm.Parameters.AddWithValue("@history", (object)data.history ?? DBNull.Value);
            cm.Parameters.AddWithValue("@cc", (object)data.cc ?? DBNull.Value);
            cm.Parameters.AddWithValue("@plan", (object)data.plan ?? DBNull.Value);
            cm.Parameters.AddWithValue("@subjective", (object)data.subjective ?? DBNull.Value);
            cm.Parameters.AddWithValue("@assessment", (object)data.assessment ?? DBNull.Value);
            cm.Parameters.AddWithValue("@ie_id", (object)data.ie_id ?? DBNull.Value);
            cm.Parameters.AddWithValue("@fu_id", (object)data.fu_id ?? DBNull.Value);
            cm.Parameters.AddWithValue("@length", (object)data.length ?? DBNull.Value);
            cm.Parameters.AddWithValue("@intake_id", (object)data.intake_id ?? DBNull.Value);
            cm.Parameters.AddWithValue("@cmp_id", (object)data.cmp_id ?? DBNull.Value);
            cm.Parameters.AddWithValue("@date", (object)data.date ?? DBNull.Value);
            cm.Parameters.AddWithValue("@file_path", (object)data.file_path ?? DBNull.Value);

            // Missing parameters added below:
            cm.Parameters.AddWithValue("@ActivitiesOfDailyLiving", (object)data.ActivitiesOfDailyLiving ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Allergies", (object)data.Allergies ?? DBNull.Value);
            cm.Parameters.AddWithValue("@CareRestrictions", (object)data.CareRestrictions ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Causation", (object)data.Causation ?? DBNull.Value);
            cm.Parameters.AddWithValue("@CoordinationExamination", (object)data.CoordinationExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Diagnoses", (object)data.Diagnoses ?? DBNull.Value);
            cm.Parameters.AddWithValue("@DiagnosticStudies", (object)data.DiagnosticStudies ?? DBNull.Value);
            cm.Parameters.AddWithValue("@FollowUp", (object)data.FollowUp ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Gait", (object)data.Gait ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Goals", (object)data.Goals ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Medications", (object)data.Medications ?? DBNull.Value);
            cm.Parameters.AddWithValue("@MotorExamination", (object)data.MotorExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@MusculoskeletalExamination", (object)data.MusculoskeletalExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@NeurologicalExamination", (object)data.NeurologicalExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@PastMedicalHistory", (object)data.PastMedicalHistory ?? DBNull.Value);
            cm.Parameters.AddWithValue("@PastSurgicalHistory", (object)data.PastSurgicalHistory ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Occupation", (object)data.Occupation ?? DBNull.Value);
            cm.Parameters.AddWithValue("@PatientInformation", (object)data.PatientInformation ?? (object)data.PatientInformation ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Precautions", (object)data.Precautions ?? DBNull.Value);
            cm.Parameters.AddWithValue("@ProceduresPerformed", (object)data.ProceduresPerformed ?? DBNull.Value);
            cm.Parameters.AddWithValue("@RangeOfMotion", (object)data.RangeOfMotion ?? DBNull.Value);
            cm.Parameters.AddWithValue("@Recommendations", (object)data.Recommendations ?? DBNull.Value);
            cm.Parameters.AddWithValue("@ReflexExamination", (object)data.ReflexExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@ReviewOfSystems", (object)data.ReviewOfSystems ?? DBNull.Value);
            cm.Parameters.AddWithValue("@SensoryExamination", (object)data.SensoryExamination ?? DBNull.Value);
            cm.Parameters.AddWithValue("@SocialHistory", (object)data.SocialHistory ?? DBNull.Value);
            cm.Parameters.AddWithValue("@VitalSigns", (object)data.VitalSigns ?? DBNull.Value);
            cm.Parameters.AddWithValue("@WorkAndAdministrativeDetails", (object)data.WorkAndAdministrativeDetails ?? DBNull.Value);

            var result = ExecuteScalar(cm);
            return Convert.ToInt32(result);
        }

        public tbl_transcribe? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_transcribe where intake_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_transcribe>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
    }
}
