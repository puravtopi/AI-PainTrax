using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUPostService : ParentService
{
    public List<tbl_post> GetAll()
    {
        List<tbl_post> dataList = ConvertDataTable<tbl_post>(GetData("select * from tbl_post"));
        return dataList;
    }

    public tbl_post? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_post where PatientFU_ID=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_post>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    

    public void Insert(tbl_post data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_post(PatientIE_ID,PatientFU_ID,CreationDate,Patient_Name,DOB,DOS,chkLeftShoulder,chkRightShoulder,chkLeftHip,chkRightHip,chkLeftKnee,chkRightKnee,chkLeftAnkleFoot,chkRightAnkleFoot,txtHistoryPresentillness,txtPhysicalExamination,txtExaminedResult)Values
				(@PatientIE_ID,@PatientFU_ID,@CreationDate,@Patient_Name,@DOB,@DOS,@chkLeftShoulder,@chkRightShoulder,@chkLeftHip,@chkRightHip,@chkLeftKnee,@chkRightKnee,@chkLeftAnkleFoot,@chkRightAnkleFoot,@txtHistoryPresentillness,@txtPhysicalExamination,@txtExaminedResult)", conn);

        cm.Parameters.AddWithValue("@PatientIE_ID", data.PatientIE_ID);
        cm.Parameters.AddWithValue("@PatientFU_ID", data.PatientFU_ID);
        cm.Parameters.AddWithValue("@CreationDate", data.CreationDate);
        cm.Parameters.AddWithValue("@Patient_Name", data.Patient_Name);
        cm.Parameters.AddWithValue("@DOB", data.DOB);
        cm.Parameters.AddWithValue("@DOS", data.DOS);
        cm.Parameters.AddWithValue("@chkLeftShoulder", data.chkLeftShoulder);
        cm.Parameters.AddWithValue("@chkRightShoulder", data.chkRightShoulder);
        cm.Parameters.AddWithValue("@chkLeftHip", data.chkLeftHip);
        cm.Parameters.AddWithValue("@chkRightHip", data.chkRightHip);
        cm.Parameters.AddWithValue("@chkLeftKnee", data.chkLeftKnee);
        cm.Parameters.AddWithValue("@chkRightKnee", data.chkRightKnee);
        cm.Parameters.AddWithValue("@chkLeftAnkleFoot", data.chkLeftAnkleFoot);
        cm.Parameters.AddWithValue("@chkRightAnkleFoot", data.chkRightAnkleFoot);
        cm.Parameters.AddWithValue("@txtHistoryPresentillness", data.txtHistoryPresentillness);       
        cm.Parameters.AddWithValue("@txtPhysicalExamination", data.txtPhysicalExamination);       
        cm.Parameters.AddWithValue("@txtExaminedResult", data.txtExaminedResult);
       
        Execute(cm);
       
    }
    public void Update(tbl_post data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_post SET
		
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
            txtPhysicalExamination=@txtPhysicalExamination,            
            txtExaminedResult=@txtExaminedResult
           
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
        cm.Parameters.AddWithValue("@txtPhysicalExamination", data.txtPhysicalExamination);       
        cm.Parameters.AddWithValue("@txtExaminedResult", data.txtExaminedResult);
        
        Execute(cm);
    
    }
    public void Delete(tbl_post data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_post
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}