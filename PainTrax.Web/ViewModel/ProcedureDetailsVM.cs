namespace PainTrax.Web.ViewModel
{
    public class ProcedureDetailsVM
    {
        public int? ProcedureDetailID { get; set; }

        public int? ProcedureDetail_ID { get; set; }
        public int? ProcedureMasterID { get; set; }
        public int? ProcedureID { get; set; }
        public int? PatientIEID { get; set; }
        public int? PatientFuID { get; set; }
        public DateTime? Consider { get; set; }
        public DateTime? Requested { get; set; }
        public DateTime? Scheduled { get; set; }
        public DateTime? Followup { get; set; }
        public DateTime? Executed { get; set; }
        public string? BodyPart { get; set; }
        public string? CreatedBy { get; set; }
        public string? Medication { get; set; }
        public string? Muscle { get; set; }
        public string? Level { get; set; }
        public string? Req_Pos { get; set; }
        public string? Sch_Pos { get; set; }
        public string? Exe_Pos { get; set; }
        public string? FU_Pos { get; set; }
        public int? IsFromNew { get; set; }
        public int? PatientProceduresID { get; set; }
        public string? Category { get; set; }
        public bool? IsConsidered { get; set; }
        public string? Side { get; set; }
        public string? Subcode { get; set; }
        public string? SignPath { get; set; }
        public string? BodypartSide { get; set; }
        public string? MC_Type { get; set; }
        public DateTime? MC_Date { get; set; }
        public string? MC_Report_Status { get; set; }
        public string? MC_Note { get; set; }
        public DateTime? MC_ReSche_Date { get; set; }
        public DateTime? CT_AUTH_Date { get; set; }
        public DateTime? CT_ReSche_Date { get; set; }
        public string? CT_Report_Status { get; set; }
        public string? CT_Note { get; set; }
        public bool? Ins_Ver_Status { get; set; }
        public string? Backup_Line { get; set; }
        public string? Ins_Note { get; set; }
        public bool? IsVaccinated { get; set; }
        public string? Vac_Status { get; set; }
        public string? Vac_Note { get; set; }
        public int? Cmp_Id { get; set; }

        public string? ADesc { get; set; }

        public string? S_ADesc { get; set; }

        public string? E_ADesc { get; set; }
    }

    public class ProcedureDetailsVMNew
    {
        public string ProcedureDetailID { get; set; }
        public string ProcedureMasterID { get; set; }
        public string PatientIEID { get; set; }
        public string PatientFuID { get; set; }
        public string SubProcedureID { get; set; }
        public string BodyPart { get; set; }
        public string ProcedureID { get; set; }
        public string Medication { get; set; }
        public string Muscle { get; set; }
        public string Level { get; set; }
        public string Position { get; set; }
        public string pocdate { get; set; }
        public string req { get; set; }
        public string BodyPartID { get; set; }
        public string IsFromNew { get; set; }
        public string PatientProcedureID { get; set; }
        public string IsConsidered { get; set; }
        public string Side { get; set; }
        public string BlobStr { get; set; }
        public string MC_Type { get; set; }
        public string MC_Date { get; set; }
        public string MC_Report_Status { get; set; }
        public string MC_Note { get; set; }
        public string MC_ReSche_Date { get; set; }
        public string CT_AUTH_Date { get; set; }
        public string CT_Report_Status { get; set; }
        public string CT_Note { get; set; }
        public string CT_ReSche_Date { get; set; }
        public string Ins_Ver_Status { get; set; }
        public string Backup_Line { get; set; }
        public string Ins_Note { get; set; }
        public string IsVaccinated { get; set; }
        public string Vac_Status { get; set; }
        public string Vac_Note { get; set; }
    }
}
