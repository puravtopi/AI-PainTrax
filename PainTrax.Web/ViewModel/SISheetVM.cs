namespace PainTrax.Web.ViewModel
{
    public class SISheetVM
    {
        public string? doe { get; set; }
        public string? Location_Id { get; set; }

    }

    public class PatientsByDOE
    {
        public Int64 IEID { get; set; }
        public Int64 FUID { get; set; }
        public string? fname { get; set; }
        public string? lname { get; set; }
        public string? state { get; set; }
        public string? city { get; set; }
        public string? type { get; set; }
        public DateTime? doa { get; set; }
        public DateTime? doe { get; set; }
        public string? compensation { get; set; }
        public string  location { get; set; }

        public List<PatientsByDOE> lstPatientsByDOE { get; set; }

    }

    public class DWSIReportVM
    {
        public int? IEID { get; set; }
        public int? FUID { get; set; }

        public string? fname { get; set; }
        public string? lname { get; set; }
        public string? VISITIEFU { get; set; }
        public string? FollowUpDate { get; set; }
        public string? doa { get; set; }
        public string? doe { get; set; }
        public string? CaseType { get; set; }
        public string? location { get; set; }
        public string? Requested { get; set; }
        public string? Scheduled { get; set; }
        public string? Alert { get; set; }
        public string? InHouse { get; set; }

        public List<DWSIReportVM> lstDWSIReport { get; set; }


    }

}
