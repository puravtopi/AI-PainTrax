using System.Data;
using System.Data.SqlClient;
//using MySql.Data.MySqlClient;
//using Microsoft.Extensions.Configuration;
using PainTrax.Web.ViewModel;
using PainTrax.Web.Models;
//using System.Reflection;
//using DocumentFormat.OpenXml.Drawing.Charts;
using MS.Services;
namespace PainTrax.Web.Services
{
    public class DataTransferService
    {
        private string _sqlServerConn;

        private readonly ProcedureService _procedureService = new ProcedureService();
        private readonly PatientIEService _patientIEService = new PatientIEService();
        private readonly PatientFUService _patientFUService = new PatientFUService();
        private readonly InscosService _inscosService = new InscosService();
        private readonly POCServices _pocServices = new POCServices();
        private readonly LogService _logServices = new LogService();



        public DataTransferService()
        {
            try
            {
                //_sqlServerConn = config.GetConnectionString("SqlServerConn");
            }
            catch (Exception ex)
            {
                tbl_log log = new tbl_log()
                {
                    CreatedBy = 1,
                    CreatedDate = System.DateTime.Now,
                    Message = ex.Message
                };
                _logServices.Insert(log);
            }



        }

        public int TransferEmployees()
        {
            int insertedCount = 0;


            try
            {


                var pdVM = new DataSet();

                //// Step 1: Read from SQL Server
                _sqlServerConn = "Data Source=10.10.93.20\\SQLEXPRESS,18667;Initial Catalog=dbPainTrax_BHF_V;uid=PTU_BHFPC;pwd=Il0ve$ql@321";
                using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                {
                    sqlConn.Open();
                    string sqlQuery = "select  p.MCODE,pd.* from tblProceduresDetail pd inner join tblProcedures p on pd.Procedure_Master_ID=p.Procedure_ID where ProcedureDetail_ID>208594";

                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    DataSet dataSet = new DataSet();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                    sqlDataAdapter.Fill(dataSet);

                    pdVM = dataSet;

                }

                foreach (DataRow row in pdVM.Tables[0].Rows)
                {

                    string cnd = " and cmp_id=18 and mcode='" + row["MCODE"] + "'";

                    var ProcedureMasterID = _procedureService.GetId(cnd);

                    cnd = " and cmp_id=18 and old_id='" + row["PatientIE_ID"] + "'";

                    var PatientIEID = _patientIEService.GetId(cnd);

                    cnd = " and cmp_id=18 and old_id='" + row["PatientFU_ID"] + "'";

                    var PatientFUID = _patientFUService.GetId(cnd);

                    ProcedureDetailsVM vm = new ProcedureDetailsVM
                    {
                        BodyPart = row["BodyPart"] != DBNull.Value ? row["BodyPart"].ToString() : string.Empty,
                        BodypartSide = string.Empty,
                        Backup_Line = string.Empty,
                        Category = string.Empty,
                        Cmp_Id = 18,
                        Consider = row["Consider"] != DBNull.Value ? Convert.ToDateTime(row["Consider"].ToString()) : null,
                        CreatedBy = string.Empty,
                        CT_AUTH_Date = null,
                        CT_Note = string.Empty,
                        CT_Report_Status = string.Empty,
                        CT_ReSche_Date = null,
                        Executed = row["Executed"] != DBNull.Value ? Convert.ToDateTime(row["Executed"].ToString()) : null,
                        Exe_Pos = row["Executed_Position"] != DBNull.Value ? row["Executed_Position"].ToString() : null,
                        Followup = row["Followup"] != DBNull.Value ? Convert.ToDateTime(row["Followup"].ToString()) : null,
                        FU_Pos = row["FU_Pos"] != DBNull.Value ? row["FU_Pos"].ToString() : string.Empty,
                        Ins_Note = null,
                        Ins_Ver_Status = null,
                        IsConsidered = row["IsConsidered"] != DBNull.Value ? Convert.ToBoolean(row["IsConsidered"]) : null,
                        IsFromNew = null,
                        IsVaccinated = null,
                        Level = row["Level"] != DBNull.Value ? row["Level"].ToString() : string.Empty,
                        MC_Date = null,
                        MC_Note = null,
                        MC_Report_Status = null,
                        MC_ReSche_Date = null,
                        Medication = row["Medication"] != DBNull.Value ? row["Medication"].ToString() : string.Empty,
                        MC_Type = null,
                        Muscle = row["Muscle"] != DBNull.Value ? row["Muscle"].ToString() : string.Empty,
                        PatientFuID = PatientFUID,
                        PatientIEID = PatientIEID,
                        PatientProceduresID = null,
                        ProcedureDetailID = null,
                        ProcedureDetail_ID = null,
                        ProcedureID = ProcedureMasterID,
                        ProcedureMasterID = ProcedureMasterID,
                        Requested = row["Requested"] != DBNull.Value ? Convert.ToDateTime(row["Requested"].ToString()) : null,
                        Req_Pos = row["Requested_Position"] != DBNull.Value ? row["Requested_Position"].ToString() : null,
                        Scheduled = row["Scheduled"] != DBNull.Value ? Convert.ToDateTime(row["Scheduled"].ToString()) : null,
                        Sch_Pos = row["Scheduled_Position"] != DBNull.Value ? row["Scheduled_Position"].ToString() : null,
                        Side = row["Sides"] != DBNull.Value ? row["Sides"].ToString() : null,
                        SignPath = row["SignPath"] != DBNull.Value ? row["SignPath"].ToString() : null,
                        Subcode = row["Subcode"] != DBNull.Value ? row["Subcode"].ToString() : null,
                        Vac_Note = null,
                        Vac_Status = null,
                        ADesc = row["ADesc"] != DBNull.Value ? row["ADesc"].ToString() : null,
                        E_ADesc = row["E_ADesc"] != DBNull.Value ? row["E_ADesc"].ToString() : null,
                        S_ADesc = row["S_ADesc"] != DBNull.Value ? row["S_ADesc"].ToString() : null,

                    };
                    _pocServices.SaveProcedureDetailsBHF(vm);
                }
            }
            catch (Exception ex)
            {
                tbl_log log = new tbl_log()
                {
                    CreatedBy = 1,
                    CreatedDate = System.DateTime.Now,
                    Message = ex.Message
                };
                _logServices.Insert(log);
            }

            return insertedCount;
        }


        public int UpdateDetails()
        {
            int insertedCount = 0;


            try
            {


                var pdVM = new DataSet();

                //// Step 1: Read from SQL Server
                _sqlServerConn = "Data Source=10.10.93.20\\SQLEXPRESS,18667;Initial Catalog=dbPainTrax_BHF_V;uid=PTU_BHFPC;pwd=Il0ve$ql@321";
                //_sqlServerConn = "Data Source=10.10.93.20\\SQLEXPRESS,18667;Initial Catalog=dbPainTrax_AKS_Live;uid=PTU_ASMPC;pwd=Il0ve$ql@321";



                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select ie.POCSummary,ie.Patient_ID from tblPatientIE as ie where (ie.POCSummary is not null and ie.POCSummary<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateProcedurePerformedOldId(row["Patient_ID"].ToString(), row["POCSummary"].ToString());
                //}


                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select ie.InsNote,ie.Patient_ID from tblPatientIE as ie where (ie.InsNote is not null and ie.InsNote<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateNoteOldId(row["Patient_ID"].ToString(), row["InsNote"].ToString());
                //}

                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select p.Patient_ID,p.MC,p.Note from tblPatientMaster p where (p.mc is not null and p.mc<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateMCNoteOldId(row["Patient_ID"].ToString(), row["MC"].ToString(), row["Note"].ToString());
                //}

                var data = _patientIEService.GetFalseMCIds(" and cmp_id=18 AND old_id IS NOT null");

                for (var i = 0; i < data.Rows.Count; i++)
                {
                    using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                    {
                        sqlConn.Open();
                        string sqlQuery = "select mc from tblPatientMaster where Patient_ID=" + data.Rows[i]["old_id"];

                        SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                        DataSet dataSet = new DataSet();
                        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                        sqlDataAdapter.Fill(dataSet);
                        sqlConn.Close();

                        if (dataSet != null)
                        {
                            if (dataSet.Tables[0].Rows.Count > 0)
                                _patientIEService.UpdateMCNoteOldId(data.Rows[i]["old_id"].ToString(), dataSet.Tables[0].Rows[0]["MC"].ToString(), "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tbl_log log = new tbl_log()
                {
                    CreatedBy = 1,
                    CreatedDate = System.DateTime.Now,
                    Message = ex.Message
                };
                _logServices.Insert(log);
            }

            return insertedCount;
        }

        public int UpdateInsurance()
        {
            int insertedCount = 0;


            try
            {


                var pdVM = new DataSet();

                //// Step 1: Read from SQL Server
                _sqlServerConn = "Data Source=10.10.93.20\\SQLEXPRESS,18667;Initial Catalog=dbPainTrax_AKS_Live;uid=PTU_BHFPC;pwd=Il0ve$ql@321";
                //_sqlServerConn = "Data Source=10.10.93.20\\SQLEXPRESS,18667;Initial Catalog=dbPainTrax_AKS_Live;uid=PTU_ASMPC;pwd=Il0ve$ql@321";



                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select ie.POCSummary,ie.Patient_ID from tblPatientIE as ie where (ie.POCSummary is not null and ie.POCSummary<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateProcedurePerformedOldId(row["Patient_ID"].ToString(), row["POCSummary"].ToString());
                //}


                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select ie.InsNote,ie.Patient_ID from tblPatientIE as ie where (ie.InsNote is not null and ie.InsNote<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateNoteOldId(row["Patient_ID"].ToString(), row["InsNote"].ToString());
                //}

                //using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                //{
                //    sqlConn.Open();
                //    string sqlQuery = "select p.Patient_ID,p.MC,p.Note from tblPatientMaster p where (p.mc is not null and p.mc<>'')";

                //    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                //    DataSet dataSet = new DataSet();
                //    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                //    sqlDataAdapter.Fill(dataSet);

                //    pdVM = dataSet;

                //}

                //foreach (DataRow row in pdVM.Tables[0].Rows)
                //{
                //    _patientIEService.UpdateMCNoteOldId(row["Patient_ID"].ToString(), row["MC"].ToString(), row["Note"].ToString());
                //}

                var data = _inscosService.GetData("  SELECT id,old_id FROM tbl_patient_ie ie WHERE ie.cmp_id=13  AND  (ie.primary_ins_cmp_id=NULL OR ie.primary_ins_cmp_id=0)");

                for (var i = 0; i < data.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(data.Rows[i]["old_id"].ToString()))
                    {
                        using (SqlConnection sqlConn = new SqlConnection(_sqlServerConn))
                        {
                            sqlConn.Open();
                            string sqlQuery = "select PatientIE_ID,insco from View_PatientIE where PatientIE_ID=" + data.Rows[i]["old_id"];

                            SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                            DataSet dataSet = new DataSet();
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                            sqlDataAdapter.Fill(dataSet);
                            sqlConn.Close();

                            if (dataSet != null)
                            {
                                if (dataSet.Tables[0].Rows.Count > 0)
                                {
                                    if (!string.IsNullOrEmpty(dataSet.Tables[0].Rows[0]["insco"].ToString()))
                                    {
                                       var result= _inscosService.GetData(" SELECT inc.id FROM tbl_inscos inc WHERE inc.cmpname='"+ dataSet.Tables[0].Rows[0]["insco"].ToString() + "' AND inc.cmp_id=13");

                                        for (var j = 0; j < result.Rows.Count; j++)
                                        {
                                            if (!string.IsNullOrEmpty(result.Rows[j]["id"].ToString()))
                                            {
                                                _patientIEService.UpdateInsurance(data.Rows[i]["id"].ToString(), result.Rows[j]["id"].ToString());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tbl_log log = new tbl_log()
                {
                    CreatedBy = 1,
                    CreatedDate = System.DateTime.Now,
                    Message = ex.Message
                };
                _logServices.Insert(log);
            }

            return insertedCount;
        }

    }
}
