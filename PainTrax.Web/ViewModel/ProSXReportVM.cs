using iText.Signatures;
using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class ProSXReportVM
    {
        public int procedureDetail_id { get; set; }
        public string name { get; set; }
        public string gender { get; set; }
        public string mc { get; set; }
        public string casetype { get; set; }
        public string location { get; set; }
        public string vaccinated { get; set; }
        public string mcode { get; set; }

        public string Phone { get; set; }
        public DateTime DOB { get; set; }
        public string ClaimNumber { get; set; }
        public string Insurance { get; set; }
        public string Allergies { get; set; }

        public DateTime? scheduled { get; set; }

        //extra columns added as per charse and arun.  
        public string? sx_center_name { get; set; }
        public string? sx_Status { get; set; }
        public string? color { get; set; }
        public string? sx_Notes { get; set; }
        public string? SX_Ins_Ver_Status { get; set; }
        public string? Ver_comment { get; set; }
        public string? Preop_notesent { get; set; }
        public string? Bookingsheet_sent { get; set; }
        public string? mc_Status { get; set; }

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<ProSXReportVM> lstProSXReport { get; set; }

        public int? locationid { get; set; }


    }
}
