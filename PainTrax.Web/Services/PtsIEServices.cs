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
    public class PtsIEServices : ParentService
    {
        //#region local variables
        //PatientIEService patientIEService = new PatientIEService();
        //FUPage1Service page1FUService = new FUPage1Service();
        //#endregion


        public List<PtsIEReportVM> GetPtsIEReport(string cnd)
        {
            string query = "";

            if (!string.IsNullOrEmpty(cnd))
            {
                query = "SELECT ie.id, CONCAT( ie.lname ,\",\", ie.fname ) AS PName,\r\nie.mobile,ie.location,ie.compensation AS \"CaseType\" ,CONVERT(ie.doe,CHAR(100)) As doe, CONVERT(ie.doa,CHAR(100)) As doa ,(Select ins.cmpname from tbl_inscos  ins where ie.primary_ins_cmp_id = ins.id) AS Ins,ie.primary_policy_no\r\n,(Select aty.Attorney from tbl_attorneys  aty where ie.attorney_id = aty.Id) AS Attorney \r\n,IFNULL((SELECT CONVERT(IFNULL(fu.doe,\"\"),CHAR(100)) FROM tbl_patient_fu fu WHERE fu.patientIE_ID = ie.id ORDER BY fu.id desc LIMIT 1 ), \"\") As LastVisit\r\n FROM vm_patient_ie ie WHERE ie.doe BETWEEN " + cnd;

            }

            //if (!string.IsNullOrEmpty(cmpid))
            //{
            //    query = query + " " + cmpid;
            //}

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = ConvertDataTable<PtsIEReportVM>(GetData(cm));
            return datalist;
        }

    }
}
