using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class PtsIEReportVM
    {
        //   public DateTime? doe { get; set; }

        public Int64 id { get; set; }
        public string? PName { get; set; }
        public string? mobile { get; set; }
        public string? location { get; set; }
        public string? CaseType { get; set; }
        public string? doe { get; set; }
        public string? doa { get; set; }
        public string? Ins { get; set; }
        public string? primary_policy_no { get; set; }
        public string? Attorney { get; set; }
        public string? LastVisit { get; set; }

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<PtsIEReportVM> lstPtsIEReport { get; set; }


    }
}
