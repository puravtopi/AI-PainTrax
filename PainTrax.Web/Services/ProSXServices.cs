using PainTrax.Services;
using MS.Services;
using PainTrax.Web.ViewModel;
using MS.Models;
using MySql.Data.MySqlClient;
using System.Data;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;

using DocumentFormat.OpenXml.Math;
using Humanizer;
using Microsoft.CodeAnalysis;
using Org.BouncyCastle.Asn1.Ocsp;

namespace PainTrax.Web.Services
{
    public class ProSXServices : ParentService
    {
        #region local variables
        PatientIEService patientIEService = new PatientIEService();
        FUPage1Service page1FUService = new FUPage1Service();
        #endregion


        public List<ProSXReportVM> GetProSXReport(string cnd)
        {
            // string query = "SELECT DISTINCT tp.ProcedureDetail_ID, CONCAT(pm.lname,' ',pm.fname)as 'Name',pm.dob as DOB,pm.mobile as Phone,ie.primary_claim_no AS ClaimNumber,(SELECT ins.cmpname FROM tbl_inscos ins WHERE ins.id =ie.primary_ins_cmp_id) AS Insurance,p1.allergies AS Allergies,pm.MC AS MC,ie.Compensation AS 'CaseType' ," +
            //"lc.location,CASE when pm.Vaccinated = 1 THEN 'Yes' ELSE 'No' END AS Vaccinated,tp.MCODE, " +
            // "CASE when pm.gender = '1' THEN 'M' when pm.gender = '2' then 'F' when pm.gender = '3' then 'O'  ELSE '' END AS gender," +
            // "tp.sx_center_name,tp.sx_Notes AS sx_Notes,tp.sx_Status,poc.color,tp.mc_Status,tp.SX_Ins_Ver_Status,tp.Ver_comment,tp.Preop_notesent,tp.Bookingsheet_sent," +
            //"tp.Scheduled  FROM tbl_Procedures_Details tp" +
            //" inner join tbl_patient_ie ie on tp.PatientIE_ID = ie.id" +
            //" inner join tbl_Patient pm on pm.id = ie.Patient_ID" +
            //" inner join tbl_locations lc ON ie.Location_ID = lc.id" +
            //" inner join tbl_attorneys a on a.id = ie.attorney_id" +
            //" left JOIN tbl_ie_page1 p1 ON p1.ie_id = ie.id AND p1.cmp_id = ie.cmp_id  AND p1.patient_id = ie.patient_id " +
            //" LEFT JOIN tbl_insurance_status_type ins ON ins.Name = tp.SX_Ins_Ver_Status " +
            //" LEFT JOIN tbl_poc_status_type poc ON poc.Name=tp.sx_Status ";

            string query = "SELECT DISTINCT tp.ProcedureDetail_ID, CONCAT(pm.lname,' ',pm.fname)as 'Name',pm.dob as DOB,pm.mobile as Phone,ie.primary_claim_no AS ClaimNumber,(SELECT ins.cmpname FROM tbl_inscos ins WHERE ins.id =ie.primary_ins_cmp_id) AS Insurance,p1.allergies AS Allergies,pm.MC AS MC,ie.Compensation AS 'CaseType' ," +
           "lc.location,CASE when pm.Vaccinated = 1 THEN 'Yes' ELSE 'No' END AS Vaccinated,tp.MCODE, " +
            "CASE when pm.gender = '1' THEN 'M' when pm.gender = '2' then 'F' when pm.gender = '3' then 'O'  ELSE '' END AS gender," +
            "tp.sx_center_name,tp.sx_Notes AS sx_Notes,tp.sx_Status,poc.color,tp.mc_Status,tp.SX_Ins_Ver_Status,tp.Ver_comment,tp.Preop_notesent,tp.Bookingsheet_sent," +
           "tp.Scheduled  FROM tbl_Procedures_Details tp" +
           " inner join tbl_patient_ie ie on tp.PatientIE_ID = ie.id" +
           " INNER JOIN tbl_Procedures pp ON pp.id = tp.Procedure_Master_ID" +
           " INNER JOIN tbl_ie_page1 p1 ON p1.ie_id = ie.id " +
           " inner join tbl_Patient pm on pm.id = ie.Patient_ID" +
           " inner join tbl_locations lc ON ie.Location_ID = lc.id" +
           " LEFT join tbl_attorneys a on a.id = ie.attorney_id" +           
           " LEFT JOIN tbl_insurance_status_type ins ON ins.Name = tp.SX_Ins_Ver_Status " +
           " LEFT JOIN tbl_poc_status_type poc ON poc.Name=tp.sx_Status ";

            if (!string.IsNullOrEmpty(cnd))
            {
                query = query + " " + cnd;
            }

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = ConvertDataTable<ProSXReportVM>(GetData(cm));
            return datalist;
        }

        public List<string> GetProSXReportDate(string cmpid)
        {
            string query = "SELECT DISTINCT DATE_FORMAT(pd.Scheduled, '%m/%d/%Y') AS Scheduled FROM tbl_procedures_details pd WHERE pd.Scheduled IS NOT NULL and pd.cmp_id=" + cmpid + " AND pd.Scheduled > CURDATE() ORDER BY pd.Scheduled desc";

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = GetData(cm);


            List<string> list = datalist.AsEnumerable()
                             .Select(row => row["Scheduled"].ToString())
                             .ToList();

            return list;
        }

        public void Update(ProSXReportVM model)
        {
            //          MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_diagcodes SET
            //BodyPart=@BodyPart,
            //DiagCode=@DiagCode,
            //Description=@Description,
            //display_order=@display_order,
            //   PreSelect=@PreSelect
            //   where Id=@Id", conn);
            //          cm.Parameters.AddWithValue("@Id", data.Id);
            //          cm.Parameters.AddWithValue("@BodyPart", data.BodyPart);
            //          cm.Parameters.AddWithValue("@DiagCode", data.DiagCode);
            //          cm.Parameters.AddWithValue("@Description", data.Description);
            //          cm.Parameters.AddWithValue("@PreSelect", data.PreSelect);
            //          cm.Parameters.AddWithValue("@display_order", data.display_order);

            //          Execute(cm);


            foreach (var item in model.lstProSXReport)
            {

                //string query = @"UPDATE tbl_Procedures_Details 
                //         SET sx_center_name = @sx_center_name,
                //             sx_Notes = @sx_Notes,
                //             sx_Status = @sx_Status,
                //             mc_Status = @mc_Status,
                //             SX_Ins_Ver_Status = @SX_Ins_Ver_Status,
                //             Ver_comment = @Ver_comment,
                //             Preop_notesent = @Preop_notesent,
                //             Bookingsheet_sent = @Bookingsheet_sent,
                //             is_booked = CASE 
                //            WHEN @Bookingsheet_sent = 'Yes' THEN 1
                //            WHEN @Bookingsheet_sent = 'No' THEN 0
                //            WHEN @Bookingsheet_sent = 'N/A' THEN 0
                //            ELSE 0
                //            ELSE 0
                //         END
                //         WHERE procedureDetail_id = @Id";

                string query = @"
            UPDATE tbl_Procedures_Details 
            SET sx_center_name = @sx_center_name,
                sx_Notes = @sx_Notes,
                sx_Status = @sx_Status,
                mc_Status = @mc_Status,
                SX_Ins_Ver_Status = @SX_Ins_Ver_Status,
                Ver_comment = @Ver_comment,
                Preop_notesent = @Preop_notesent,
                Bookingsheet_sent = @Bookingsheet_sent,
                is_booked = CASE 
                                WHEN @Bookingsheet_sent = 'Yes' THEN 1
                                WHEN @Bookingsheet_sent IN ('No', 'N/A') THEN 0
                                ELSE 0
                            END
            WHERE procedureDetail_id = @Id";
                // using (var con = new MySqlConnection(ConfigurationManager.ConnectionStrings["MySqlConn"].ConnectionString))
                using (var cmd = new MySqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Id", item.procedureDetail_id);
                    cmd.Parameters.AddWithValue("@sx_center_name", item.sx_center_name ?? "");
                    cmd.Parameters.AddWithValue("@sx_Notes", item.sx_Notes ?? "");
                    cmd.Parameters.AddWithValue("@sx_Status", item.sx_Status ?? "");
                    cmd.Parameters.AddWithValue("@mc_Status", item.mc_Status ?? "");
                    cmd.Parameters.AddWithValue("@SX_Ins_Ver_Status", item.SX_Ins_Ver_Status ?? "");
                    cmd.Parameters.AddWithValue("@Ver_comment", item.Ver_comment ?? "");
                    cmd.Parameters.AddWithValue("@Preop_notesent", item.Preop_notesent ?? "");
                    cmd.Parameters.AddWithValue("@Bookingsheet_sent", item.Bookingsheet_sent ?? "");

                    //conn.Open();
                    // cmd.ExecuteNonQuery();
                    Execute(cmd);
                }


            }



        }

    }
}
