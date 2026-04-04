using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class SignatureService : ParentService
    {
        public List<tbl_signature> GetAll(string cnd = "")
        {

            List<tbl_signature> dataList = null;


            string query = "select * from tbl_signature where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;


            dataList = ConvertDataTable<tbl_signature>(GetData(query));
            return dataList;
        }
        public tbl_signature? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_signature where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_signature>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int Insert(tbl_signature data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_signature
        (fname,lname,dob,signaturePath,cmp_id) VALUES
            (@fname,@lname,@dob,@signaturePath,@cmp_id); select @@identity", conn);
            cm.Parameters.AddWithValue("@fname", data.fname);
            cm.Parameters.AddWithValue("@lname", data.lname);
            cm.Parameters.AddWithValue("@dob", data.dob);
            cm.Parameters.AddWithValue("@signaturePath", data.signaturePath);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            var result = Execute(cm);
            return result;

        }

        public void Update(tbl_signature data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_signature SET 
                id=@id,   
                fname=@fname,
                lname=@lname,
                dob=@dob,
				signaturePath=@signaturePath,
                cmp_id=@cmp_id where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@fname", data.fname);
            cm.Parameters.AddWithValue("@lname", data.lname);
            cm.Parameters.AddWithValue("@dob", data.dob);
            cm.Parameters.AddWithValue("@signaturePath", data.signaturePath); ;
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);

            Execute(cm);
        }
        public void Delete(tbl_signature data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE from tbl_signature where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }

        public int ManageSign(tbl_ie_sign data)
        {

            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_ie_sign where patient_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.patient_id);
            var datalist = ConvertDataTable<tbl_signature>(GetData(cm)).FirstOrDefault();

            int Result = 0;

            if (datalist != null)
            {
                cm = new MySqlCommand(@"UPDATE tbl_ie_sign SET
				signatureData=@signatureData where patient_id=@patient_id", conn);
                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@signatureData", data.signatureData);

                Execute(cm);
                Result = 1;
            }
            else
            {
                cm = new MySqlCommand(@"INSERT INTO tbl_ie_sign
        (patient_id, signatureData,signatureValue) VALUES
            (@patient_id, @signatureData,@signatureValue); select @@identity", conn);

                cm.Parameters.AddWithValue("@patient_id", data.patient_id);
                cm.Parameters.AddWithValue("@signatureData", data.signatureData);
                cm.Parameters.AddWithValue("@signatureValue", null);
                var result = ExecuteScalar(cm);
                Result = Convert.ToInt32(result);
            }
            return Result;

        }
    }
}
