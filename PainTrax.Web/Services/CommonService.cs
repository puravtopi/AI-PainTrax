using MS.Models;
using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class CommonService : ParentService
    {
        public List<tbl_pages> GetPagesAll(string cnd = "")
        {
            string query = "select * from tbl_pages where 1=1 ";
            query += cnd;

            List<tbl_pages> dataList = ConvertDataTable<tbl_pages>(GetData(query));
            return dataList;
        }

        public List<tbl_reports> GetReportsAll(string cnd = "")
        {
            string query = "select * from tbl_reports where 1=1 ";
            query += cnd;

            List<tbl_reports> dataList = ConvertDataTable<tbl_reports>(GetData(query));
            return dataList;
        }

        public List<tbl_role> GetRoleAll(string cnd = "")
        {
            string query = "select * from tbl_role where 1=1 ";
            query += cnd;

            List<tbl_role> dataList = ConvertDataTable<tbl_role>(GetData(query));
            return dataList;
        }

        public List<tbl_treatment_master> GetTreatmentAll(string cnd = "")
        {
            string query = "select * from tbl_treatment_master where 1=1 ";

            query += cnd;
            List<tbl_treatment_master> dataList = ConvertDataTable<tbl_treatment_master>(GetData(query));
            return dataList;
        }

        public void UpdateAccountNo(string patientId, string accountNo)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_patient SET account_no=@accountNo WHERE id=@id", conn);
            cm.Parameters.AddWithValue("@id", patientId);
            cm.Parameters.AddWithValue("@accountNo", accountNo);

            Execute(cm);
        }

        public PatientDetails getPatientDetails(string id)
        {

            string query = "SELECT fname,lname,dob FROM vm_patient_ie where patient_id=@id";

            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<PatientDetails>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int InsertSign(string signData, string signValue, string patient_id)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_patient_authoackno
        (signData,signValue,patient_id,signDate) VALUES
            (@signData,@signValue,@patient_id,@signDate); select @@identity;", conn);

            cm.Parameters.AddWithValue("@signData", signData);
            cm.Parameters.AddWithValue("@signValue", signValue);
            cm.Parameters.AddWithValue("@patient_id", patient_id);
         
            cm.Parameters.AddWithValue("@signDate", System.DateTime.Now.ToString("yyyy-MM-dd"));
            var result = ExecuteScalar(cm);
            return Convert.ToInt32(result);
        }

        public bool isSignExist(string id)
        {

            string query = "SELECT id FROM tbl_patient_authoackno where patient_id=@id";

            
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = GetData(cm);
            if (datalist != null)
            {
                if (datalist.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

        public bool deleteAuthoSign(string id)
        {

            try
            {
                string query = "delete from tbl_patient_authoackno where patient_id=@id";

                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand(query, conn);
                cm.Parameters.AddWithValue("@id", id);
                
                Execute(cm);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public PatientDetails getPatientDetailsWithSign(string id)
        {

            string query = "SELECT p.fname,p.lname,p.dob,pa.signData as sign FROM tbl_patient_authoackno pa JOIN tbl_patient p ON pa.patient_id=p.id where p.id=@id";

            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<PatientDetails>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
    }
}
