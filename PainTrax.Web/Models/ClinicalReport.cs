namespace PainTrax.Web.Models
{
    public class ClinicalReport
    {
        public string History { get; set; }
        public string ChiefComplaint { get; set; }
        public string Subjective { get; set; }
        public string Assessment { get; set; }
        public string Plan { get; set; }
        public string Transcript { get; set; }
    }

    public class ReportRequest
    {
        public int length { get; set; }
        public int intake_id { get; set; }
        public string Transcript { get; set; }
        public string file_path { get; set; }
    }

    public class TranscribeData
    {
        public string PatientName { get; set; }
        public string DOE { get; set; }
        public string Transcript { get; set; }
        public string FilePath { get; set; }
        public int Length { get; set; } = 0;
    }
}
