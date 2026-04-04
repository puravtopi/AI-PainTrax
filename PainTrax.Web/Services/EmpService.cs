using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class EmpService : ParentService
{
    public List<tbl_emp> GetAll(string cnd="")
    {
        string query = "select * from tbl_emp where 1=1 ";
        query += cnd;
        List<tbl_emp> dataList = ConvertDataTable<tbl_emp>(GetData(query));
        return dataList;
    }

    public tbl_emp? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_emp where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_emp>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public int Insert(tbl_emp data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_emp
		(name,address,phone,fax,patient_id)Values
				(@name,@address,@phone,@fax,@patient_id);select @@identity;", conn);
        cm.Parameters.AddWithValue("@name", data.name);
        cm.Parameters.AddWithValue("@address", data.address);
        cm.Parameters.AddWithValue("@phone", data.phone);
        cm.Parameters.AddWithValue("@fax", data.fax);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        var result = ExecuteScalar(cm);
        return result;
    }
    public void Update(tbl_emp data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_emp SET
		name=@name,
		address=@address,
		phone=@phone,
		fax=@fax,
		patient_id=@patient_id		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@name", data.name);
        cm.Parameters.AddWithValue("@address", data.address);
        cm.Parameters.AddWithValue("@phone", data.phone);
        cm.Parameters.AddWithValue("@fax", data.fax);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        Execute(cm);
    }
    public void Delete(tbl_emp data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_emp
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}