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
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace PainTrax.Web.Services
{
    public class ProSXDetailsServices : ParentService
    {
        //#region local variables
        //PatientIEService patientIEService = new PatientIEService();
        //FUPage1Service page1FUService = new FUPage1Service();
        //#endregion


        public List<ProSXDetailsReportVM> GetPtsIEReport(string cnd)
        {
            string query = "";

            if (!string.IsNullOrEmpty(cnd))
            {
                query = "SELECT a.Attorney,ie.doe AS DOE,ie.primary_policy_no AS PolicyNo,pd.ProcedureDetail_ID,CASE WHEN ie.gender =1 THEN 'Male' ELSE 'FeMale' END AS SEX,CONCAT(ie.lname ,\", \",ie.fname) AS NAME,ie.MC AS MC,ie.Compensation as 'CaseType' ,ie.location,CASE when ie.Vaccinated = 1 THEN 'Yes' ELSE 'No' END AS Vaccinated,pd.MCODE,pd.BodyPart, CASE when pd.Ins_Ver_Status=1 THEN 'YES' ELSE 'No' END as Ins_ver_status , CASE when pd.MC_TYPE='Yes' then IFNULL(pd.MC_Report_Status,'') ELSE 'Received' end as MC_Status , CASE when ie.Compensation='WC' then IFNULL(pd.CT_Report_Status,'Received') else 'Received' end as 'Case_Status' , CASE when pd.Ins_Ver_Status=1 then IFNULL(pd.Backup_Line,'Received') else 'Received' end as 'InsVerStatus'  , CASE when pd.IsVaccinated=1 then IFNULL(pd.Vac_Status,'Received') else 'Received' end as 'Vac_Status' ,CONVERT( IFNULL(DATE(pd.Scheduled),''),CHAR(15)) as Scheduled ,CONVERT( IFNULL(DATE(pd.Executed),''),CHAR(15)) as Executed ,CONVERT( IFNULL(DATE(pd.Requested),''),CHAR(15)) as Requested    \r\nFROM vm_patient_ie ie INNER JOIN tbl_procedures_details pd ON pd.PatientIE_ID = ie.id INNER JOIN tbl_attorneys a ON a.Id = ie.attorney_id" + cnd;

            }

            //if (!string.IsNullOrEmpty(cmpid))
            //{
            //    query = query + " " + cmpid;
            //}

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = ConvertDataTable<ProSXDetailsReportVM>(GetData(cm));
            return datalist;
        }

    }
}
