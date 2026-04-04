using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUPreService : ParentService
{
    public List<tbl_pre> GetAll()
    {
        List<tbl_pre> dataList = ConvertDataTable<tbl_pre>(GetData("select * from tbl_pre"));
        return dataList;
    }

    public tbl_pre? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_pre where PatientFU_ID=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_pre>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public tbl_preOPDefault? GetOneOPDefault(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("SELECT ie.compensation AS casetype,a.pmh AS txtPastMedicalHistory ,a.psh as txtpastsurgicalhistory,a.medication AS txtdailyMedications,a.allergies AS txtAllergies,a.family_history AS txtFamilyHistory,a.social_history AS txtSH, a.history AS txtPresentillness FROM tbl_patient_ie ie LEFT JOIN  tbl_ie_page1 a ON ie.patient_id = a.patient_id AND ie.id = a.ie_id AND ie.patient_id = a.patient_id WHERE ie.id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_preOPDefault>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public void Insert(tbl_pre data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_pre
		(PatientIE_ID,PatientFU_ID,CreationDate,Patient_Name,DOB,DOS,chkLeftShoulder,chkRightShoulder,chkLeftHip,chkRightHip,chkLeftKnee,chkRightKnee,chkLeftAnkleFoot,chkRightAnkleFoot,txtHistoryPresentillness,txtpresentcomplain,txtPastMedicalHistory,txtpastsurgicalhistory,txtdailyMedications,txtAllergies,txtpastaccideninjuries,txtSocialHistory,txtPhysicalExamination,txtDiagnosticImaging,txtAssestmentplan,txtExaminedResult,txtDefault,txtNote,txtPresentillness,txtFamilyHistory,txtSH)Values
				(@PatientIE_ID,@PatientFU_ID,@CreationDate,@Patient_Name,@DOB,@DOS,@chkLeftShoulder,@chkRightShoulder,@chkLeftHip,@chkRightHip,@chkLeftKnee,@chkRightKnee,@chkLeftAnkleFoot,@chkRightAnkleFoot,@txtHistoryPresentillness,@txtpresentcomplain,@txtPastMedicalHistory,@txtpastsurgicalhistory,@txtdailyMedications,@txtAllergies,@txtpastaccideninjuries,@txtSocialHistory,@txtPhysicalExamination,@txtDiagnosticImaging,@txtAssestmentplan,@txtExaminedResult,@txtDefault,@txtNote,@txtPresentillness,@txtFamilyHistory,@txtSH)", conn);

        cm.Parameters.AddWithValue("@PatientIE_ID", data.PatientIE_ID);
        cm.Parameters.AddWithValue("@PatientFU_ID", data.PatientFU_ID);
        cm.Parameters.AddWithValue("@CreationDate", data.CreationDate);
        cm.Parameters.AddWithValue("@Patient_Name", data.Patient_Name);
        cm.Parameters.AddWithValue("@DOB", data.DOB);
        cm.Parameters.AddWithValue("@DOS", data.DOS);
        cm.Parameters.AddWithValue("@chkLeftShoulder", data.chkLeftShoulder);
        cm.Parameters.AddWithValue("@chkRightShoulder", data.chkRightShoulder);
        cm.Parameters.AddWithValue("@chkLeftHip", data.chkLeftHip);
        cm.Parameters.AddWithValue("@chkRightHip", data.chkRightHip); ;
        cm.Parameters.AddWithValue("@chkLeftKnee", data.chkLeftKnee);
        cm.Parameters.AddWithValue("@chkRightKnee", data.chkRightKnee);
        cm.Parameters.AddWithValue("@chkLeftAnkleFoot", data.chkLeftAnkleFoot);
        cm.Parameters.AddWithValue("@chkRightAnkleFoot", data.chkRightAnkleFoot);
        cm.Parameters.AddWithValue("@txtHistoryPresentillness", data.txtHistoryPresentillness);
        cm.Parameters.AddWithValue("@txtpresentcomplain", data.txtpresentcomplain);
        cm.Parameters.AddWithValue("@txtPastMedicalHistory", data.txtPastMedicalHistory);
        cm.Parameters.AddWithValue("@txtpastsurgicalhistory", data.txtpastsurgicalhistory);
        cm.Parameters.AddWithValue("@txtdailyMedications", data.txtdailyMedications);
        cm.Parameters.AddWithValue("@txtAllergies", data.txtAllergies);
        cm.Parameters.AddWithValue("@txtpastaccideninjuries", data.txtpastaccideninjuries);
        cm.Parameters.AddWithValue("@txtSocialHistory", data.txtSocialHistory);
        cm.Parameters.AddWithValue("@txtPhysicalExamination", data.txtPhysicalExamination);
        cm.Parameters.AddWithValue("@txtDiagnosticImaging", data.txtDiagnosticImaging);
        cm.Parameters.AddWithValue("@txtAssestmentplan", data.txtAssestmentplan);
        cm.Parameters.AddWithValue("@txtExaminedResult", data.txtExaminedResult);
        cm.Parameters.AddWithValue("@txtDefault", data.txtDefault);
        cm.Parameters.AddWithValue("@txtNote", data.txtNote);
        cm.Parameters.AddWithValue("@txtPresentillness", data.txtPresentillness);
        cm.Parameters.AddWithValue("@txtFamilyHistory", data.txtFamilyHistory);
        cm.Parameters.AddWithValue("@txtSH", data.txtSH);
        Execute(cm);
        //var result = Execute(cm);
        //cm.ExecuteNonQuery();
        //Execute(cm);
    }
    public void Update(tbl_pre data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_pre SET
		
            PatientIE_ID=@PatientIE_ID,
            PatientFU_ID=@PatientFU_ID,
            CreationDate=@CreationDate,
            Patient_Name=@Patient_Name,
            DOB=@DOB,
            DOS=@DOS,
            chkLeftShoulder=@chkLeftShoulder,
            chkRightShoulder=@chkRightShoulder,
            chkLeftHip=@chkLeftHip,
            chkRightHip=@chkRightHip,
            chkLeftKnee=@chkLeftKnee,
            chkRightKnee=@chkRightKnee,
            chkLeftAnkleFoot=@chkLeftAnkleFoot,
            chkRightAnkleFoot=@chkRightAnkleFoot,
            txtHistoryPresentillness=@txtHistoryPresentillness,
            txtpresentcomplain=@txtpresentcomplain,
            txtPastMedicalHistory=@txtPastMedicalHistory,
            txtpastsurgicalhistory=@txtpastsurgicalhistory,
            txtdailyMedications=@txtdailyMedications,
            txtAllergies=@txtAllergies,
            txtpastaccideninjuries=@txtpastaccideninjuries,
            txtSocialHistory=@txtSocialHistory,
            txtPhysicalExamination=@txtPhysicalExamination,
            txtDiagnosticImaging=@txtDiagnosticImaging,
            txtAssestmentplan=@txtAssestmentplan,
            txtExaminedResult=@txtExaminedResult,
            txtDefault=@txtDefault,
            txtNote=@txtNote,
            txtPresentillness=@txtPresentillness,
            txtFamilyHistory=@txtFamilyHistory,
            txtSH=@txtSH
        where id=@id", conn);

        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@PatientIE_ID", data.PatientIE_ID);
        cm.Parameters.AddWithValue("@PatientFU_ID", data.PatientFU_ID);
        cm.Parameters.AddWithValue("@CreationDate", data.CreationDate);
        cm.Parameters.AddWithValue("@Patient_Name", data.Patient_Name);
        cm.Parameters.AddWithValue("@DOB", data.DOB);
        cm.Parameters.AddWithValue("@DOS", data.DOS);
        cm.Parameters.AddWithValue("@chkLeftShoulder", data.chkLeftShoulder);
        cm.Parameters.AddWithValue("@chkRightShoulder", data.chkRightShoulder);
        cm.Parameters.AddWithValue("@chkLeftHip", data.chkLeftHip);
        cm.Parameters.AddWithValue("@chkRightHip", data.chkRightHip); ;
        cm.Parameters.AddWithValue("@chkLeftKnee", data.chkLeftKnee);
        cm.Parameters.AddWithValue("@chkRightKnee", data.chkRightKnee);
        cm.Parameters.AddWithValue("@chkLeftAnkleFoot", data.chkLeftAnkleFoot);
        cm.Parameters.AddWithValue("@chkRightAnkleFoot", data.chkRightAnkleFoot);
        cm.Parameters.AddWithValue("@txtHistoryPresentillness", data.txtHistoryPresentillness);
        cm.Parameters.AddWithValue("@txtpresentcomplain", data.txtpresentcomplain);
        cm.Parameters.AddWithValue("@txtPastMedicalHistory", data.txtPastMedicalHistory);
        cm.Parameters.AddWithValue("@txtpastsurgicalhistory", data.txtpastsurgicalhistory);
        cm.Parameters.AddWithValue("@txtdailyMedications", data.txtdailyMedications);
        cm.Parameters.AddWithValue("@txtAllergies", data.txtAllergies);
        cm.Parameters.AddWithValue("@txtpastaccideninjuries", data.txtpastaccideninjuries);
        cm.Parameters.AddWithValue("@txtSocialHistory", data.txtSocialHistory);
        cm.Parameters.AddWithValue("@txtPhysicalExamination", data.txtPhysicalExamination);
        cm.Parameters.AddWithValue("@txtDiagnosticImaging", data.txtDiagnosticImaging);
        cm.Parameters.AddWithValue("@txtAssestmentplan", data.txtAssestmentplan);
        cm.Parameters.AddWithValue("@txtExaminedResult", data.txtExaminedResult);
        cm.Parameters.AddWithValue("@txtDefault", data.txtDefault);
        cm.Parameters.AddWithValue("@txtNote", data.txtNote);
        cm.Parameters.AddWithValue("@txtPresentillness", data.txtPresentillness);
        cm.Parameters.AddWithValue("@txtFamilyHistory", data.txtFamilyHistory);
        cm.Parameters.AddWithValue("@txtSH", data.txtSH);
        Execute(cm);
    
    }
    public void Delete(tbl_pre data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_pre
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}