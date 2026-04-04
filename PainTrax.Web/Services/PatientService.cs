using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace MS.Services;
public class PatientService : ParentService
{
    public List<tbl_patient> GetAll(string cnd="")
    {

        List<tbl_patient> dataList = null;


        string query = "select * from tbl_patient where 1=1 ";

        if (!string.IsNullOrEmpty(cnd))
            query = query + cnd;


        dataList = ConvertDataTable<tbl_patient>(GetData(query));
        return dataList;
    }

    public tbl_patient? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_patient where id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_patient>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public List<tbl_patient> GetAautoComplete(string cnd = "")
    {
        string query = "select *,state as label,state as val from tbl_patient where 1=1 ";
        query += cnd;
        List<tbl_patient> dataList = ConvertDataTable<tbl_patient>(GetData(query));
        return dataList;
    }
    public int Insert(tbl_patient data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_patient
		(fname,lname,mname,gender,dob,age,email,handeness,ssn,address,city,state,zip,home_ph,mobile,vaccinated,mc,mc_details,account_no,createddate,createdby,cmp_id)Values
				(@fname,@lname,@mname,@gender,@dob,@age,@email,@handeness,@ssn,@address,@city,@state,@zip,@home_ph,@mobile,@vaccinated,@mc,@mc_details,@account_no,@createddate,@createdby,@cmp_id);select @@identity;", conn);
        cm.Parameters.AddWithValue("@fname", data.fname);
        cm.Parameters.AddWithValue("@lname", data.lname);
        cm.Parameters.AddWithValue("@mname", data.mname);
        cm.Parameters.AddWithValue("@gender", data.gender);
        cm.Parameters.AddWithValue("@dob", data.dob);
        cm.Parameters.AddWithValue("@age", data.age);
        cm.Parameters.AddWithValue("@email", data.email);
        cm.Parameters.AddWithValue("@handeness", data.handeness);
        cm.Parameters.AddWithValue("@ssn", data.ssn);
        cm.Parameters.AddWithValue("@address", data.address);
        cm.Parameters.AddWithValue("@city", data.city);
        cm.Parameters.AddWithValue("@state", data.state);
        cm.Parameters.AddWithValue("@zip", data.zip);
        cm.Parameters.AddWithValue("@home_ph", data.home_ph);
        cm.Parameters.AddWithValue("@mobile", data.mobile);
        cm.Parameters.AddWithValue("@vaccinated", data.vaccinated);
        cm.Parameters.AddWithValue("@mc", data.mc);
        cm.Parameters.AddWithValue("@account_no", data.account_no);
        cm.Parameters.AddWithValue("@createddate", System.DateTime.Now);
        cm.Parameters.AddWithValue("@createdby", data.createdby);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        //cm.Parameters.AddWithValue("@physicianid", data.physicianid);
        cm.Parameters.AddWithValue("@mc_details", data.mc_details);
        var result = ExecuteScalar(cm);
        return result;
    }
    public void Update(tbl_patient data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_patient SET
		fname=@fname,
		lname=@lname,
		mname=@mname,
		gender=@gender,
		dob=@dob,
		age=@age,
		email=@email,
		handeness=@handeness,
		ssn=@ssn,
		address=@address,
		city=@city,
		state=@state,
		zip=@zip,
		home_ph=@home_ph,
		mobile=@mobile,
		vaccinated=@vaccinated,
		mc=@mc,
        mc_details=@mc_details,
		account_no=@account_no,        
		updatedate=@updatedate,
		updatedby=@updatedby


		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        cm.Parameters.AddWithValue("@fname", data.fname);
        cm.Parameters.AddWithValue("@lname", data.lname);
        cm.Parameters.AddWithValue("@mname", data.mname);
        cm.Parameters.AddWithValue("@gender", data.gender);
        cm.Parameters.AddWithValue("@dob", data.dob);
        cm.Parameters.AddWithValue("@age", data.age);
        cm.Parameters.AddWithValue("@email", data.email);
        cm.Parameters.AddWithValue("@handeness", data.handeness);
        cm.Parameters.AddWithValue("@ssn", data.ssn);
        cm.Parameters.AddWithValue("@address", data.address);
        cm.Parameters.AddWithValue("@city", data.city);
        cm.Parameters.AddWithValue("@state", data.state);
        cm.Parameters.AddWithValue("@zip", data.zip);
        cm.Parameters.AddWithValue("@home_ph", data.home_ph);
        cm.Parameters.AddWithValue("@mobile", data.mobile);
        cm.Parameters.AddWithValue("@vaccinated", data.vaccinated);
        cm.Parameters.AddWithValue("@mc", data.mc);
        cm.Parameters.AddWithValue("@mc_details", data.mc_details);
        cm.Parameters.AddWithValue("@account_no", data.account_no);
       // cm.Parameters.AddWithValue("@physicianid", data.physicianid);
        cm.Parameters.AddWithValue("@updatedate", System.DateTime.Now);
        cm.Parameters.AddWithValue("@updatedby", data.updatedby);

        Execute(cm);
    }
    public void Delete(tbl_patient data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_patient
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

    public List<string> GetPatientSearchList(int cmp_id, string prefix)
    {
        string cnd = " and cmp_id=" + cmp_id + " and (fname like  '" + prefix + "%' or lname like '" + prefix + "%')";
        var data = GetAll(cnd);
        List<string> _patients = new List<string>();

        foreach (var item in data)
        {
            _patients.Add(item.fname + " " + item.lname + "_" + item.id.ToString());
        }

        return _patients;
    }

    

}