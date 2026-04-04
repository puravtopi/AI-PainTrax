namespace PainTrax.Web.ViewModel
{
    public class POCVM
    {
        public int? ProcedureId { get; set; }
        public string MCODE { get; set; }
        public int? INhouseProc { get; set; }
        public string HasPosition { get; set; }
        public int? HasLevel { get; set; }
        public int? HasMuscle { get; set; }
        public int? HasSubCode { get; set; }
        public int? BID { get; set; }
        public string Bodypart { get; set; }
        public int? HasMedication { get; set; }
        public int? PatientProceduresID { get; set; }
        public string Medication { get; set; }

        public string Muscle { get; set; }
        public string Level { get; set; }
        public DateTime? Consider { get; set; }
        public DateTime? Requested { get; set; }
        public DateTime? Scheduled { get; set; }
        public DateTime? Executed { get; set; }
        public DateTime? Followup { get; set; }
        public string Req_Pos { get; set; }
        public string Sch_Pos { get; set; }
        public string FU_Pos { get; set; }
        public string Exe_Pos { get; set; }
        public string SubProcedureID { get; set; }
        public int? PatientIEID { get; set; }
        public int? PatientFuID { get; set; }
        public int? Count { get; set; }
        public string Sides { get; set; }
        public int? HasSides { get; set; }
        public int? Display_Order { get; set; }
        public int? ProcedureDetailID { get; set; }
        public string LevelsDefault { get; set; }
        public string SidesDefault { get; set; }
        public string SignPath { get; set; }
        public int? CF { get; set; }
    }
}
