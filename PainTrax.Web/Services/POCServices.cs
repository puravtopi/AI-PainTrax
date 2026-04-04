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
    public class POCServices : ParentService
    {
        #region local variables
        PatientIEService patientIEService = new PatientIEService();
        FUPage1Service page1FUService = new FUPage1Service();
        #endregion

        public string[] GetInjuredParts(int patientIE, int patientFU = 0)
        {
            try
            {

                string[] result = null;

                if (patientFU == 0)
                {
                    var data = patientIEService.GetOnePage1(patientIE);


                    if (data != null)
                    {
                        if (data.bodypart != null)
                            result = data.bodypart.Split(',');
                    }
                }
                else
                {
                    var data = page1FUService.GetOne(patientFU);


                    if (data != null)
                    {
                        if (data.bodypart != null)
                            result = data.bodypart.Split(',');
                    }
                }

                return result;

            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public string[] GetInjuredPartsPOC(int patientIE)
        {
            try
            {

                string[] result = null;


                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("CALL sp_injur_body_parts(" + patientIE + ")", conn);

                var datalist = GetData(cm);


                if (datalist != null)
                {
                    if (datalist.Rows.Count > 0)
                    {
                        var bodypart = datalist.Rows[0][0].ToString();

                        if (bodypart != null)
                        {
                            string[] val = bodypart.Split(',');

                            result = val
           .Select(s => new { Original = s, Normalized = NormalizeString(s) }) // Normalize strings
           .GroupBy(x => x.Normalized) // Group by normalized value
           .Select(g => g.First().Original) // Select the first original string from each group
           .ToArray();
                        }
                    }
                }


                return result;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string NormalizeString(string input)
        {
            return new string(input.Where(c => !char.IsWhiteSpace(c)).ToArray()) // Remove spaces
                .ToLower(); // Convert to lowercase
        }

        public System.Data.DataTable GetAllProcedures(string bodyParts, int patientIEID, string potion, int cmp_id)
        {
            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("CALL sp_GetAllProceduress('" + bodyParts + "'," + patientIEID + ",'" + potion + "'," + cmp_id + ")", conn);

                var datalist = GetData(cm);

                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public System.Data.DataTable GetAllProceduresFU(string bodyParts, int patientFUID, string potion, int cmp_id)
        {
            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("CALL sp_GetAllProcedures_FU('" + bodyParts + "','" + potion + "'," + patientFUID + "," + cmp_id + ")", conn);

                var datalist = GetData(cm);

                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public string GetMuscle(int patientIEID)
        {
            string strresult = "";
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select  HasMuscle from tbl_Procedures where id =@Procedure_ID ", conn);
            cm.Parameters.AddWithValue("@Procedure_ID", patientIEID);
            var result = GetData(cm);

            if (result != null && result.Rows.Count > 0)
            {
                strresult = result.Rows[0]["HasMuscle"].ToString();
            }

            return strresult;
        }

        public string GetSubCode(int patientIEID)
        {
            string strresult = "";
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select HasSubProcedure from tbl_Procedures where id=@Procedure_ID ", conn);
            cm.Parameters.AddWithValue("@Procedure_ID", patientIEID);
            var result = GetData(cm);

            if (result != null && result.Rows.Count > 0)
            {
                strresult = result.Rows[0]["HasSubProcedure"].ToString();
            }

            return strresult;
        }

        public string GetMedication(int patientIEID)
        {
            string strresult = "";
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select HasMedication from tbl_Procedures where id=@Procedure_ID ", conn);
            cm.Parameters.AddWithValue("@Procedure_ID", patientIEID);
            var result = GetData(cm);

            if (result != null && result.Rows.Count > 0)
            {
                strresult = result.Rows[0]["HasMedication"].ToString();
            }

            return strresult;
        }

        public string SaveProcedureDetails(ProcedureDetailsVM model)
        {
            List<MySqlParameter> param = new List<MySqlParameter>();

            param.Add(new MySqlParameter("v_Backup_Line", model.Backup_Line));
            param.Add(new MySqlParameter("v_BodyPart", model.BodyPart));
            param.Add(new MySqlParameter("v_BodypartSide", model.BodypartSide));
            param.Add(new MySqlParameter("v_Category", model.Category));
            param.Add(new MySqlParameter("v_Consider", model.Consider));
            param.Add(new MySqlParameter("v_CreatedBy", model.CreatedBy));
            param.Add(new MySqlParameter("v_CT_AUTH_Date", model.CT_AUTH_Date));
            param.Add(new MySqlParameter("v_CT_Note", model.CT_Note));
            param.Add(new MySqlParameter("v_CT_Report_Status", model.CT_Report_Status));
            param.Add(new MySqlParameter("v_CT_ReSche_Date", model.CT_ReSche_Date));
            param.Add(new MySqlParameter("v_Executed", model.Executed));
            param.Add(new MySqlParameter("v_Exe_Pos", model.Exe_Pos));
            param.Add(new MySqlParameter("v_Followup", model.Followup));
            param.Add(new MySqlParameter("v_FU_Pos", model.FU_Pos));
            param.Add(new MySqlParameter("v_Ins_Note", model.Ins_Note));
            param.Add(new MySqlParameter("v_Ins_Ver_Status", model.Ins_Ver_Status));
            param.Add(new MySqlParameter("v_IsConsidered", model.IsConsidered));
            param.Add(new MySqlParameter("v_IsFromNew", model.IsFromNew));
            param.Add(new MySqlParameter("v_IsVaccinated", model.IsVaccinated));
            param.Add(new MySqlParameter("v_Level", model.Level));
            param.Add(new MySqlParameter("v_MC_Date", model.MC_Date));
            param.Add(new MySqlParameter("v_MC_Note", model.MC_Note));
            param.Add(new MySqlParameter("v_MC_Report_Status", model.MC_Report_Status));
            param.Add(new MySqlParameter("v_MC_ReSche_Date", model.MC_ReSche_Date));
            param.Add(new MySqlParameter("v_MC_Type", model.MC_Type));
            param.Add(new MySqlParameter("v_Medication", model.Medication));
            param.Add(new MySqlParameter("v_Muscle", model.Muscle));
            param.Add(new MySqlParameter("v_PatientFuID", model.PatientFuID));
            param.Add(new MySqlParameter("v_PatientIEID", model.PatientIEID));
            param.Add(new MySqlParameter("v_PatientProceduresID", model.PatientProceduresID));
            param.Add(new MySqlParameter("v_ProcedureDetailID", model.ProcedureDetailID));
            param.Add(new MySqlParameter("v_ProcedureID", model.ProcedureID));
            param.Add(new MySqlParameter("v_ProcedureMasterID", model.ProcedureMasterID));
            param.Add(new MySqlParameter("v_Requested", model.Requested));
            param.Add(new MySqlParameter("v_Req_Pos", model.Req_Pos));
            param.Add(new MySqlParameter("v_Scheduled", model.Scheduled));
            param.Add(new MySqlParameter("v_Sch_Pos", model.Sch_Pos));
            param.Add(new MySqlParameter("v_Side", model.Side));
            param.Add(new MySqlParameter("v_SignPath", model.SignPath));
            param.Add(new MySqlParameter("v_Subcode", model.Subcode));
            param.Add(new MySqlParameter("v_Vac_Note", model.Vac_Note));
            param.Add(new MySqlParameter("v_Vac_Status", model.Vac_Status));
            param.Add(new MySqlParameter("v_cmpid", model.Cmp_Id));

            var data = ExecuteSP(CommonSp.SavePatientProceduresDetails, param);

            return data;
        }

        public string SaveProcedureDetailsBHF(ProcedureDetailsVM model)
        {
            List<MySqlParameter> param = new List<MySqlParameter>();

            param.Add(new MySqlParameter("v_Backup_Line", model.Backup_Line));
            param.Add(new MySqlParameter("v_BodyPart", model.BodyPart));
            param.Add(new MySqlParameter("v_BodypartSide", model.BodypartSide));
            param.Add(new MySqlParameter("v_Category", model.Category));
            param.Add(new MySqlParameter("v_Consider", model.Consider));
            param.Add(new MySqlParameter("v_CreatedBy", model.CreatedBy));
            param.Add(new MySqlParameter("v_CT_AUTH_Date", model.CT_AUTH_Date));
            param.Add(new MySqlParameter("v_CT_Note", model.CT_Note));
            param.Add(new MySqlParameter("v_CT_Report_Status", model.CT_Report_Status));
            param.Add(new MySqlParameter("v_CT_ReSche_Date", model.CT_ReSche_Date));
            param.Add(new MySqlParameter("v_Executed", model.Executed));
            param.Add(new MySqlParameter("v_Exe_Pos", model.Exe_Pos));
            param.Add(new MySqlParameter("v_Followup", model.Followup));
            param.Add(new MySqlParameter("v_FU_Pos", model.FU_Pos));
            param.Add(new MySqlParameter("v_Ins_Note", model.Ins_Note));
            param.Add(new MySqlParameter("v_Ins_Ver_Status", model.Ins_Ver_Status));
            param.Add(new MySqlParameter("v_IsConsidered", model.IsConsidered));
            param.Add(new MySqlParameter("v_IsFromNew", model.IsFromNew));
            param.Add(new MySqlParameter("v_IsVaccinated", model.IsVaccinated));
            param.Add(new MySqlParameter("v_Level", model.Level));
            param.Add(new MySqlParameter("v_MC_Date", model.MC_Date));
            param.Add(new MySqlParameter("v_MC_Note", model.MC_Note));
            param.Add(new MySqlParameter("v_MC_Report_Status", model.MC_Report_Status));
            param.Add(new MySqlParameter("v_MC_ReSche_Date", model.MC_ReSche_Date));
            param.Add(new MySqlParameter("v_MC_Type", model.MC_Type));
            param.Add(new MySqlParameter("v_Medication", model.Medication));
            param.Add(new MySqlParameter("v_Muscle", model.Muscle));
            param.Add(new MySqlParameter("v_PatientFuID", model.PatientFuID));
            param.Add(new MySqlParameter("v_PatientIEID", model.PatientIEID));
            param.Add(new MySqlParameter("v_PatientProceduresID", model.PatientProceduresID));
            param.Add(new MySqlParameter("v_ProcedureDetailID", model.ProcedureDetailID));
            param.Add(new MySqlParameter("v_ProcedureID", model.ProcedureID));
            param.Add(new MySqlParameter("v_ProcedureMasterID", model.ProcedureMasterID));
            param.Add(new MySqlParameter("v_Requested", model.Requested));
            param.Add(new MySqlParameter("v_Req_Pos", model.Req_Pos));
            param.Add(new MySqlParameter("v_Scheduled", model.Scheduled));
            param.Add(new MySqlParameter("v_Sch_Pos", model.Sch_Pos));
            param.Add(new MySqlParameter("v_Side", model.Side));
            param.Add(new MySqlParameter("v_SignPath", model.SignPath));
            param.Add(new MySqlParameter("v_Subcode", model.Subcode));
            param.Add(new MySqlParameter("v_Vac_Note", model.Vac_Note));
            param.Add(new MySqlParameter("v_Vac_Status", model.Vac_Status));
            param.Add(new MySqlParameter("ADesc", model.ADesc));
            param.Add(new MySqlParameter("E_ADesc", model.E_ADesc));
            param.Add(new MySqlParameter("S_ADesc", model.S_ADesc));

            var data = ExecuteSP(CommonSp.SavePatientProceduresDetailsBHF, param);

            return data;
        }

        public ProcedureDetailsVM? GetProcedureDetails(int ProcedurrDetail_ID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetProcedureDetails(" + ProcedurrDetail_ID + ")", conn);

                var datalist = ConvertDataTable<ProcedureDetailsVM>(GetData(cm)).FirstOrDefault();
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public System.Data.DataTable GetProcedureCountDetails(int ProcedureID, int PatientIEID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetProcBycount(" + ProcedureID + "," + PatientIEID + ")", conn);

                var datalist = (GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public DataTable GetPOCIE(int PatientIEID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetPOC(" + PatientIEID + ")", conn);

                var datalist = (GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public DataTable GetPOCIENew(int PatientIEID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetPOC_IE_Print(" + PatientIEID + ")", conn);

                var datalist = (GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public DataTable GetPOCFU(int PatientFUID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetPOC_FU(" + PatientFUID + ")", conn);

                var datalist = (GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public DataTable GetPOCFUPrint(int PatientFUID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand("CALL sp_GetPOC_FU_Print(" + PatientFUID + ")", conn);

                var datalist = (GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public List<POCSumarryVM> GetExecutedPOCIE(int PatientIEID)
        {
            try
            {

                string query = "CALL sp_GetExecutedPOC(" + PatientIEID + ")";

                List<POCSumarryVM> dataList = ConvertDataTable<POCSumarryVM>(GetData(query));
                return dataList;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public List<POCSumarryVM> GetPOCSummary(int PatientIEID)
        {
            try
            {

                string query = "CALL sp_GetPOC(" + PatientIEID + ")";

                List<POCSumarryVM> dataList = ConvertDataTable<POCSumarryVM>(GetData(query));
                return dataList;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public List<POCSumarryVM> GetFUPOCSummary(int PatientFUID)
        {
            try
            {

                string query = "CALL sp_GetPOC_FU(" + PatientFUID + ")";

                List<POCSumarryVM> dataList = ConvertDataTable<POCSumarryVM>(GetData(query));
                return dataList;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public bool RemoveProcedureCountDetails(int ProcedureDetailsID)
        {
            try
            {

                MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_procedures_details where ProcedureDetail_ID=@id", conn);
                cm.Parameters.AddWithValue("@id", ProcedureDetailsID);
                Execute(cm);
                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public int ForwardPOCIETOFU(int PatientIEID, int PatientFUID)
        {
            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("CALL sp_ForwardPOC_IETOFU('" + PatientIEID + "','" + PatientFUID + "')", conn);

                var datalist = GetData(cm);

                return 1;

            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public List<POCReportVM> GetPOCReport(string cnd)
        {
            //string query = "SELECT pm.id,replace(pm.lname, ' ','') AS lname,replace(pm.fname, ' ','') AS fname, tp.ProcedureDetail_ID,CONCAT(pm.lname,', ',pm.fname)as 'Name',CASE when pm.MC=1 THEN 'Yes' ELSE 'No' END as MC," +
            //"ie.Compensation AS 'CaseType' ,ie.doa,pm.dob,ie.doe,pm.mobile AS Phone,ie.primary_policy_no,ie.primary_claim_no,ins.cmpname,tp.sides,tp.level," +
            //"lc.location,CASE when pm.Vaccinated = 1 THEN 'Yes' ELSE 'No' END AS Vaccinated,tp.MCODE ,CONCAT(u.fname,' ',u.lname) as providerName, " +
            //"tp.Requested,p1.allergies,p1.note," +
            //"tp.surgercy_center,tp.surgon_name,tp.assistent_name,sc.Surgerycenter_name," +
            //"tp.Executed," +
            //"CASE when pm.gender = '1' THEN 'Male' when pm.gender = '2' then 'Female' when pm.gender = '3' then 'Other'  ELSE '' END AS gender," +
            //"tp.Scheduled,tp.sx_center_name FROM tbl_Procedures_Details tp" +
            //" inner join tbl_patient_ie ie on tp.PatientIE_ID = ie.id" +
            //" inner join tbl_Procedures pp on pp.id=tp.Procedure_Master_ID" +
            //" INNER JOIN tbl_ie_page1 p1 ON ie.id = p1.ie_id " +
            //" inner join tbl_Patient pm on pm.id = ie.Patient_ID" +
            //" inner join tbl_locations lc ON ie.Location_ID = lc.id" +
            //" LEFT JOIN tbl_inscos ins ON ie.primary_ins_cmp_id = ins.id" +
            // " LEFT JOIN tbl_surgerycenter sc ON tp.surgercy_center = sc.Id  " +
            //" LEFT JOIN tbl_users u ON ie.provider_id = u.id";
            // query += " where  (tp.Scheduled>='" + DateTime.Now.Date.ToString("yyyy/MM/dd") + "')";


            string query = "SELECT a.Attorney,pm.id,REPLACE(pm.lname,' ','') AS lname,REPLACE(pm.fname,' ','') AS fname,tp.ProcedureDetail_ID," +
                "CONCAT(pm.lname, ', ', pm.fname) AS Name,CASE WHEN pm.MC = 'TRUE' THEN 'Yes' WHEN  pm.MC = 'FALSE' THEN 'No' ELSE pm.MC END AS MC,pm.mc_details as MC_Details,pm.account_no as acct_no,ie.Compensation AS CaseType," +
                "ie.doa, pm.dob,IFNULL((SELECT _fu.doe FROM tbl_patient_fu _fu WHERE _fu.patientIE_ID=ie.id ORDER BY _fu.doe DESC LIMIT 0,1),ie.doe) AS doe, pm.mobile AS Phone,ie.primary_policy_no AS PolicyNo, ie.primary_claim_no,ins.cmpname, tp.sides, tp.level,lc.location," +
                "CASE WHEN pm.Vaccinated=1 THEN 'Yes' ELSE 'No' END AS Vaccinated,tp.MCODE,CONCAT(u.fname,' ',u.lname) AS providerName," +
                "tp.Requested, p1.allergies, pm.mc_details as note,tp.surgercy_center, tp.surgon_name, tp.assistent_name,sc.Surgerycenter_name, tp.Executed," +
                "CASE WHEN pm.gender='1' THEN 'Male' WHEN pm.gender='2' THEN 'Female' WHEN pm.gender='3' THEN 'Other' ELSE '' END AS gender," +
                "tp.Scheduled, tp.sx_center_name,ROW_NUMBER() OVER (PARTITION BY tp.ProcedureDetail_ID ORDER BY pm.id DESC) AS rn FROM tbl_Procedures_Details tp" +
                " INNER JOIN tbl_patient_ie ie ON tp.PatientIE_ID = ie.id INNER JOIN tbl_Procedures pp ON pp.id = tp.Procedure_Master_ID INNER JOIN tbl_ie_page1 p1" +
                " ON ie.id = p1.ie_id INNER JOIN tbl_Patient pm ON pm.id = ie.Patient_ID INNER JOIN tbl_locations lc ON ie.Location_ID = lc.id" +
                " LEFT JOIN tbl_inscos ins ON ie.primary_ins_cmp_id = ins.id LEFT JOIN tbl_surgerycenter sc ON tp.surgercy_center = sc.Id " +
                " LEFT JOIN tbl_users u ON ie.provider_id = u.id LEFT JOIN tbl_attorneys a ON ie.attorney_id = a.Id ";


            if (!string.IsNullOrEmpty(cnd))
            {
                query = query + " " + cnd;
            }

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = ConvertDataTable<POCReportVM>(GetData(cm));
            return datalist;
        }


        public List<SurgoryCenterDashboardVM> GetSurgoryDashboardData(string fDate, string tDate, string cmpId)
        {
            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("CALL sp_Get_Suregory_Dashboard(" + cmpId + ", '" + fDate + "','" + tDate + "')", conn);

                var datalist = ConvertDataTable<SurgoryCenterDashboardVM>(GetData(cm));
                return datalist;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void TransferToExecute(string id, string sDate)
        {
            string query = "update tbl_Procedures_Details set Executed='" + Convert.ToDateTime(sDate).ToString("yyyy/MM/dd") + "',Scheduled=null where ProcedureDetail_ID=" + id;
            MySqlCommand cm = new MySqlCommand(query, conn);

            Execute(cm);
        }

        public void TransferToReschedules(string id, string sDate = "")
        {
            string query = "update tbl_Procedures_Details set Scheduled=null where ProcedureDetail_ID=" + id;
            if (sDate != "")
                query = "update tbl_Procedures_Details set Scheduled='" + sDate + "' where ProcedureDetail_ID=" + id;

            MySqlCommand cm = new MySqlCommand(query, conn);

            Execute(cm);
        }
        public void UpdatePOCReportSideandLevel(string id, string side = "", string level = "")
        {
            string query = "";
            if (side != "")
                query = "update tbl_Procedures_Details set Sides='" + side + "',Level='" + level + "' where ProcedureDetail_ID=" + id;

            MySqlCommand cm = new MySqlCommand(query, conn);

            Execute(cm);
        }

        public void deleteMCode(string mcode)
        {
            string query = "delete from  tbl_Procedures where mcode='" + mcode + "' and cmp_id=2";
            MySqlCommand cm = new MySqlCommand(query, conn);

            Execute(cm);
        }

        public bool UpdatePOCSurgoryCenter(string sProcedureDetailIDs, string sId, string sSCName, string sAssistant, string sSurgeon, int cmpid,DateTime sDate)
        {
            try
            {
                string query = "call sp_Update_SurgeryCenter(" + sId + ",'" + sSCName + "','" + sAssistant + "', '" + sProcedureDetailIDs.TrimStart(',') + "','" + sSurgeon.TrimStart(',') + "'," + cmpid + "," +
                    "'"+ sDate.ToString("yyyy-MM-dd") +"')";

                MySqlCommand cm = new MySqlCommand(query, conn);

                Execute(cm);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
