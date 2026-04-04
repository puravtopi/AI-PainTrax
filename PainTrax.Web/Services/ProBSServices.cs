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
    public class ProBSServices : ParentService
    {
        #region local variables
        PatientIEService patientIEService = new PatientIEService();
        FUPage1Service page1FUService = new FUPage1Service();
        #endregion


        public List<ProBSReportVM> GetProBSReport(string cnd)
        {           

            string query = "SELECT DISTINCT tp.ProcedureDetail_ID, CONCAT(pm.lname,' ',pm.fname)as 'Name',pm.account_no," +
           "pm.dob as DOB,CASE when pm.gender = '1' THEN 'M' when pm.gender = '2' then 'F' when pm.gender = '3' then 'O'  ELSE '' END AS gender, " +
            "pm.mobile as Phone,pm.address AS Address,pm.city AS City,pm.state AS State,pm.zip AS Zip,pm.ssn AS SSN," +
            "ie.Compensation AS 'CaseType',IFNULL((SELECT _fu.doe FROM tbl_patient_fu _fu WHERE _fu.patientIE_ID=ie.id ORDER BY _fu.doe DESC LIMIT 0,1),ie.doe) AS doe," +
            "ie.doa,a.Attorney AS AttorneyName,a.ContactNo AS AttorneyPhone,tp.MCODE,tp.sides, tp.level,lc.location, " +
            "(SELECT ins.cmpname FROM tbl_inscos ins WHERE ins.id =ie.primary_ins_cmp_id) AS Insurance,ie.primary_claim_no AS ClaimNumber,ie.primary_wcb_group AS WCB," +
            "tp.sx_center_name,poc.color,tp.Bookingsheet_sent,pm.mc_details as Note, " +
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

            var datalist = ConvertDataTable<ProBSReportVM>(GetData(cm));
            return datalist;
        }

        

        

    }
}
