using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services;
public class MacrosMasterService : ParentService
{
    public List<tbl_macros_master> GetAll(string cnd = "")
    {
        string query = "select * from tbl_macros_master where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_macros_master> dataList = ConvertDataTable<tbl_macros_master>(GetData(query));
        return dataList;
    }

    public tbl_macros_master? GetOne(tbl_macros_master data)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_macros_master where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        var datalist = ConvertDataTable<tbl_macros_master>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public void Insert(tbl_macros_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_macros_master
		(bodypart,heading,cc_desc,pe_desc,a_desc,p_desc,cf,pn,rom_desc,pc_desc,ros_desc,ds_desc,pt_desc,drd_desc,drd_notes,pre_select,cmp_id,created_date,created_by)Values
				(@bodypart,@heading,@cc_desc,@pe_desc,@a_desc,@p_desc,@cf,@pn,@rom_desc,@pc_desc,@ros_desc,@ds_desc,@pt_desc,@drd_desc,@drd_notes,@pre_select,@cmp_id,@created_date,@created_by)", conn);
        cm.Parameters.AddWithValue("@bodypart", data.bodypart);
        cm.Parameters.AddWithValue("@heading", data.heading);
        cm.Parameters.AddWithValue("@cc_desc", data.cc_desc);
        cm.Parameters.AddWithValue("@pe_desc", data.pe_desc);
        cm.Parameters.AddWithValue("@a_desc", data.a_desc);
        cm.Parameters.AddWithValue("@p_desc", data.p_desc);
        cm.Parameters.AddWithValue("@cf", data.cf);
        cm.Parameters.AddWithValue("@pn", data.pn);
        cm.Parameters.AddWithValue("@pre_select", data.pre_select);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@created_date", data.created_date);
        cm.Parameters.AddWithValue("@created_by", data.created_by);
        cm.Parameters.AddWithValue("@rom_desc", data.rom_desc);

        cm.Parameters.AddWithValue("@pc_desc", data.pc_desc);
        cm.Parameters.AddWithValue("@ros_desc", data.ros_desc);
        cm.Parameters.AddWithValue("@ds_desc", data.ds_desc);
        cm.Parameters.AddWithValue("@pt_desc", data.pt_desc);
        cm.Parameters.AddWithValue("@drd_desc", data.drd_desc);
        cm.Parameters.AddWithValue("@drd_notes", data.drd_notes);

        Execute(cm);
    }
    public void Update(tbl_macros_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_macros_master SET
		bodypart=@bodypart,
		heading=@heading,
		cc_desc=@cc_desc,
		pe_desc=@pe_desc,
		a_desc=@a_desc,
        p_desc=@p_desc,
        rom_desc=@rom_desc,
        pc_desc=@pc_desc,
        ros_desc=@ros_desc,
        ds_desc=@ds_desc,
        pt_desc=@pt_desc,
        drd_desc=@drd_desc,
        drd_notes=@drd_notes,
		cf=@cf,
		pn=@pn,
		pre_select=@pre_select,		
		updated_date=@updated_date,
		updated_by=@updated_by		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@bodypart", data.bodypart);
        cm.Parameters.AddWithValue("@heading", data.heading);
        cm.Parameters.AddWithValue("@cc_desc", data.cc_desc);
        cm.Parameters.AddWithValue("@pe_desc", data.pe_desc);
        cm.Parameters.AddWithValue("@a_desc", data.a_desc);
        cm.Parameters.AddWithValue("@p_desc", data.p_desc);
        cm.Parameters.AddWithValue("@rom_desc", data.rom_desc);
        cm.Parameters.AddWithValue("@pc_desc", data.pc_desc);
        cm.Parameters.AddWithValue("@ros_desc", data.ros_desc);
        cm.Parameters.AddWithValue("@ds_desc", data.ds_desc);
        cm.Parameters.AddWithValue("@pt_desc", data.pt_desc);
        cm.Parameters.AddWithValue("@drd_desc", data.drd_desc);
        cm.Parameters.AddWithValue("@drd_notes", data.drd_notes);

        
        cm.Parameters.AddWithValue("@cf", data.cf);
        cm.Parameters.AddWithValue("@pn", data.pn);
        cm.Parameters.AddWithValue("@pre_select", data.pre_select);

        cm.Parameters.AddWithValue("@updated_date", data.updated_date);
        cm.Parameters.AddWithValue("@updated_by", data.updated_by);

        

        Execute(cm);
    }
    public void Delete(tbl_macros_master data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_macros_master
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}