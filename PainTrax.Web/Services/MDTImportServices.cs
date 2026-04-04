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
    public class MDTImportServices : ParentService
    {
        #region local variables
        PatientIEService patientIEService = new PatientIEService();
        FUPage1Service page1FUService = new FUPage1Service();
        #endregion


        public List<MDTImportReportVM> GeMDTImportReport(string cnd)
        {
            string query = "";

            if (!string.IsNullOrEmpty(cnd))
            {

                    query = "SELECT ie.doe,ie.id AS PatientIE_ID,pm.lname,pm.fname,pm.mname,CASE when pm.gender = '1' THEN 'Male' when pm.gender = '2' then 'Female' when pm.gender = '3' then 'Other'  ELSE '' END AS gender,pm.dob,pm.address,pm.city,pm.state,pm.zip,pm.home_ph,pm.mobile,loc.location " +
                    " from tbl_patient_ie ie INNER JOIN tbl_patient pm ON ie.patient_id = pm.id INNER JOIN tbl_locations loc ON loc.id= ie.location_id " +
                    " WHERE (ie.doe BETWEEN " + cnd + ") ";

            }

            //if (!string.IsNullOrEmpty(cmpid))
            //{
            //    query = query + " " + cmpid;
            //}

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = ConvertDataTable<MDTImportReportVM>(GetData(cm));
            return datalist;
        }

    }
}
