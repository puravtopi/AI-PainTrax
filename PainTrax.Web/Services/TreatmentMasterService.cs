using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;
using MS.Models;

namespace MS.Services;
public class TreatmentMasterService : ParentService
{
    public List<tbl_treatment_master> GetAll(string cnd = "")
    {

        string query = "select * from tbl_treatment_master where 1=1 ";
        query += cnd + " Order BY display_order ASC";
        List<tbl_treatment_master> dataList = ConvertDataTable<tbl_treatment_master>(GetData(query));
        return dataList;
    }

    public tbl_treatment_master? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_treatment_master where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_treatment_master>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public int Insert(tbl_treatment_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_treatment_master
		(treatment_details,pre_select,display_order,cmp_id)Values
	    (@treatment_details,@pre_select,@display_order,@cmp_id);select @@identity", conn);
        cm.Parameters.AddWithValue("@treatment_details", data.treatment_details);        
        cm.Parameters.AddWithValue("@pre_select", data.pre_select);
        cm.Parameters.AddWithValue("@display_order", data.display_order);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        var result = ExecuteScalar(cm);
        return result;
    }
    public void Update(tbl_treatment_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_treatment_master SET
		treatment_details=@treatment_details,pre_select=@pre_select,display_order=@display_order
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@pre_select", data.pre_select);
        cm.Parameters.AddWithValue("@treatment_details", data.treatment_details);
        cm.Parameters.AddWithValue("@display_order", data.display_order);

        Execute(cm);
    }
    public void Delete(tbl_treatment_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_treatment_master
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}