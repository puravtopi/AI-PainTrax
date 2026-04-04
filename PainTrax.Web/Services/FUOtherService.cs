using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUOtherService : ParentService
{
    public List<tbl_fu_other> GetAll()
    {
        List<tbl_fu_other> dataList = ConvertDataTable<tbl_fu_other>(GetData("select * from tbl_fu_other"));
        return dataList;
    }

    public tbl_fu_other? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_fu_other where fu_id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_fu_other>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public void Insert(tbl_fu_other data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_fu_other
		(ie_id,patient_id,treatment_details,note1,note2,note3,treatment_delimit,followup_duration,followup_date,fu_id,treatment_delimit_desc)Values
				(@ie_id,@patient_id,@treatment_details,@note1,@note2,@note3,@treatment_delimit,@followup_duration,@followup_date,@fu_id,@treatment_delimit_desc)", conn);
        cm.Parameters.AddWithValue("@ie_id", data.ie_id);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@treatment_details", data.treatment_details);
        cm.Parameters.AddWithValue("@note1", data.note1);
        cm.Parameters.AddWithValue("@note2", data.note2);
        cm.Parameters.AddWithValue("@note3", data.note3);
        cm.Parameters.AddWithValue("@treatment_delimit", data.treatment_delimit);
        cm.Parameters.AddWithValue("@followup_duration", data.followup_duration);
        cm.Parameters.AddWithValue("@treatment_delimit_desc", string.IsNullOrEmpty(data.treatment_delimit_desc) ? "" : data.treatment_delimit_desc.TrimStart('^'));
        cm.Parameters.AddWithValue("@followup_date", data.followup_date);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        Execute(cm);
    }
    public void Update(tbl_fu_other data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_other SET
		ie_id=@ie_id,
			
		patient_id=@patient_id,
		treatment_details=@treatment_details,
		note1=@note1,
		note2=@note2,
		note3=@note3,
		treatment_delimit=@treatment_delimit,
		treatment_delimit_desc=@treatment_delimit_desc,
		followup_duration=@followup_duration,
		followup_date=@followup_date where fu_id=@fu_id", conn);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        cm.Parameters.AddWithValue("@ie_id", data.ie_id);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@treatment_details", data.treatment_details);
        cm.Parameters.AddWithValue("@note1", data.note1);
        cm.Parameters.AddWithValue("@note2", data.note2);
        cm.Parameters.AddWithValue("@note3", data.note3);
        cm.Parameters.AddWithValue("@treatment_delimit", data.treatment_delimit);
        cm.Parameters.AddWithValue("@treatment_delimit_desc", data.treatment_delimit_desc == null ? "" : data.treatment_delimit_desc.TrimStart('^'));
        cm.Parameters.AddWithValue("@followup_duration", data.followup_duration);
        cm.Parameters.AddWithValue("@followup_date", data.followup_date);

        Execute(cm);
    }
    public void Delete(tbl_fu_other data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_other
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}