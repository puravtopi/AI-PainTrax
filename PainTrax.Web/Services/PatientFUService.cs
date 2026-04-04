using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using PainTrax.Web.ViewModel;
using PainTrax.Web.Models;

namespace MS.Services;
public class PatientFUService : ParentService
{
    public List<tbl_patient_fu> GetAll()
    {
        List<tbl_patient_fu> dataList = ConvertDataTable<tbl_patient_fu>(GetData("select * from tbl_patient_fu"));
        return dataList;
    }

    public tbl_patient_fu? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_patient_fu where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_patient_fu>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public DateTime? GetFirstIEDateOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select doe from tbl_patient_ie where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_patient_ie>(GetData(cm)).FirstOrDefault();
        return datalist == null ? null : datalist.doe;
    }

    public List<vm_patient_fu> GetAllByIeId(int ie_id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from vm_patient_fu where patientIE_ID=@id order by doe desc", conn);
        cm.Parameters.AddWithValue("@id", ie_id);
        var datalist = ConvertDataTable<vm_patient_fu>(GetData(cm));
        return datalist;
    }

    public int Insert(tbl_patient_fu data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_patient_fu
		(patient_id,provider_id,patientIE_ID,doe,created_date,created_by,updated_date,updated_by,is_active,cmp_id,extra_comments,type,accident_type,physicianid,location_id,procedure_performed)Values
				(@patient_id,@provider_id,@patientIE_ID,@doe,@created_date,@created_by,@updated_date,@updated_by,@is_active,@cmp_id,@extra_comments,@type,@accident_type,@physicianid,@location_id,@procedure_performed);select @@identity;", conn);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@provider_id", data.provider_id);

        cm.Parameters.AddWithValue("@patientIE_ID", data.patientIE_ID);
        cm.Parameters.AddWithValue("@doe", data.doe);
        cm.Parameters.AddWithValue("@created_date", data.created_date);
        cm.Parameters.AddWithValue("@created_by", data.created_by);
        cm.Parameters.AddWithValue("@updated_date", data.updated_date);
        cm.Parameters.AddWithValue("@updated_by", data.updated_by);
        cm.Parameters.AddWithValue("@is_active", data.is_active);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@extra_comments", data.extra_comments);
        cm.Parameters.AddWithValue("@type", data.type);
        cm.Parameters.AddWithValue("@accident_type", data.accident_type);
        cm.Parameters.AddWithValue("@physicianid", data.physicianid);
        cm.Parameters.AddWithValue("@location_id", data.location_id);
        cm.Parameters.AddWithValue("@procedure_performed", data.procedure_performed);

        var result = ExecuteScalar(cm);
        return result;
    }
    public void Update(tbl_patient_fu data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_patient_fu SET
		patient_id=@patient_id,
        provider_id=@provider_id,
		patientIE_ID=@patientIE_ID,
		doe=@doe,
		
		created_date=@created_date,
		created_by=@created_by,
		updated_date=@updated_date,
		updated_by=@updated_by,
		is_active=@is_active,
		cmp_id=@cmp_id,
		location_id=@location_id,
		extra_comments=@extra_comments,
        accident_type=@accident_type,
        physicianid=@physicianid
			where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@provider_id", data.provider_id);

        cm.Parameters.AddWithValue("@patientIE_ID", data.patientIE_ID);
        cm.Parameters.AddWithValue("@doe", data.doe);

        cm.Parameters.AddWithValue("@created_date", data.created_date);
        cm.Parameters.AddWithValue("@created_by", data.created_by);
        cm.Parameters.AddWithValue("@updated_date", data.updated_date);
        cm.Parameters.AddWithValue("@updated_by", data.updated_by);
        cm.Parameters.AddWithValue("@is_active", data.is_active);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@extra_comments", data.extra_comments);
        cm.Parameters.AddWithValue("@accident_type", data.accident_type);
        cm.Parameters.AddWithValue("@physicianid", data.physicianid);
        cm.Parameters.AddWithValue("@location_id", data.location_id);
        Execute(cm);
    }
    public void Delete(int fuId)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_patient_fu
		where id=@id;delete from tbl_procedures_details where Patientfu_Id=@id;", conn);
        cm.Parameters.AddWithValue("@id", fuId);
        Execute(cm);
    }

    public int UpdatePage1Plan(int id, string plan)
    {
        MySqlCommand cm = new MySqlCommand(@"update tbl_fu_page1 set
		plan=@plan
       
        where fu_id=@id
				 ;select 1;", conn);

        cm.Parameters.AddWithValue("@id", id);
        cm.Parameters.AddWithValue("@plan", plan);

        var result = ExecuteScalar(cm);
        return result;
    }

    public bool IsFuExist(int id,string doe, string type,int ie_id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select id from tbl_patient_fu where DATE(doe)=DATE(@doe) AND TYPE=@type AND id<>@id and patientIE_ID=@ie_id", conn);
        cm.Parameters.AddWithValue("@id", id);
        cm.Parameters.AddWithValue("@doe", doe);
        cm.Parameters.AddWithValue("@type", type);
        cm.Parameters.AddWithValue("@ie_id", ie_id);
        var datalist = ConvertDataTable<tbl_patient_fu>(GetData(cm)).FirstOrDefault();
        if (datalist == null)
            return false;
        else {
            return true;
        }
    }

    public int? GetId(string cnd = "")
    {
        string query = "SELECT * FROM tbl_patient_fu where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_patient_fu> dataList = ConvertDataTable<tbl_patient_fu>(GetData(query));
        if (dataList != null)
        {
            if (dataList.Count > 0)
            {
                return Convert.ToInt32(dataList[0].id);
            }
            else
                return null;
        }
        else
            return null;
    }


    #region POC
    public int UpdateProcedureExecuteDate(string fDate, string tDate, string fuId)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_procedures_details
SET Executed = @tDate
WHERE Executed = @fDate and PatientFU_ID=@fuId;Select 1;", conn);

        cm.Parameters.AddWithValue("@tDate", tDate);
        cm.Parameters.AddWithValue("@fDate", fDate);
        cm.Parameters.AddWithValue("@fuId", fuId);
        var result = ExecuteScalar(cm);
        return Convert.ToInt32(result);
    }

    public void UpdateProcedurePerformed(string id, string procedure_performed)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_patient_fu SET
		procedure_performed=@procedure_performed	
        where id=@id", conn);
        cm.Parameters.AddWithValue("@id", id);
        cm.Parameters.AddWithValue("@procedure_performed", procedure_performed);
        Execute(cm);
    }
    #endregion

}