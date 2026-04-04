using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace PainTrax.Web.Services
{
    public class DashboardService : ParentService
    {

        public int GetTotalPatient(int cmp_id, int type = 4)
        {

            string query = "", fDate = "", tDate = "";

            if (type == 1)
            {
                fDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                tDate = System.DateTime.Now.ToString("yyyy-MM-dd");
            }
            else if (type == 2)
            {
                DateTime now = DateTime.Now;
                fDate = new DateTime(now.Year, now.Month, 1).ToString("yyyy-MM-dd");
                tDate = new DateTime(now.Year, now.Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            }
            else if (type == 3)
            {
                int year = DateTime.Now.Year;
                fDate = new DateTime(year, 1, 1).ToString("yyyy-MM-dd");
                tDate = new DateTime(year, 12, 31).ToString("yyyy-MM-dd");
            }
            if (type == 4)
                query = "SELECT count(id) FROM tbl_patient_ie WHERE patient_id IN (SELECT id FROM tbl_patient WHERE  cmp_id=@cmp_id)";
            else
                query = "SELECT count(id) FROM tbl_patient_ie WHERE patient_id IN (SELECT id FROM tbl_patient WHERE  cmp_id=@cmp_id and createddate BETWEEN '" + fDate + "' AND '" + tDate + "')";

            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }

        public int GetTotalAttorny(int cmp_id)
        {

            string query = "SELECT count(id) FROM tbl_patient_ie WHERE attorney_id=0 AND patient_id IN (SELECT id FROM tbl_patient WHERE cmp_id=@cmp_id) ";
            
            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }
        public int GetTotalClaimNo(int cmp_id)
        {

            string query = "SELECT count(id) FROM tbl_patient_ie WHERE (primary_claim_no IS NULL OR primary_claim_no='') AND patient_id IN (SELECT id FROM tbl_patient WHERE cmp_id=@cmp_id) ";

            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }

        public int GetTotalInsuranceCompany(int cmp_id)
        {
            string query = "SELECT count(id) FROM tbl_patient_ie WHERE primary_ins_cmp_id=0 AND patient_id IN (SELECT id FROM tbl_patient WHERE cmp_id=@cmp_id) ";

            MySqlCommand cm = new MySqlCommand(query, conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);

            var result = ExecuteScalar(cm);
            return result;
        }
    }
}
