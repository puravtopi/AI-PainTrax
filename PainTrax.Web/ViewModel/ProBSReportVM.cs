using iText.Signatures;
using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class ProBSReportVM
    {
        public int procedureDetail_id { get; set; }
        public string name { get; set; }
        public string account_no { get; set; }
        public DateTime? DOB { get; set; }
        public string gender { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string SSN { get; set; }
        public string casetype { get; set; }
        public DateTime? doe { get; set; }
        public DateTime? doa { get; set; }
        public string AttorneyName { get; set; }
        public string AttorneyPhone { get; set; }
        public string mcode { get; set; }
        public string sides { get; set; }
        public string level { get; set; }
        public string location { get; set; }
        public string Insurance { get; set; }
        public string ClaimNumber { get; set; }
        public string WCB { get; set; } 
        public DateTime? scheduled { get; set; }       
        public string? sx_center_name { get; set; }        
        public string? color { get; set; }        
        public string? Bookingsheet_sent { get; set; }
        public string? Note { get; set; }
        public string? Procedures { get; set; }
        public string? cptcodes { get; set; }
        public string? icdcodes { get; set; }
        public string? specialequ { get; set; }       

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<ProBSReportVM> lstProBSReport { get; set; }

        public int? locationid { get; set; }


    }
}
