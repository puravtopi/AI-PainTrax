using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class POCReportVM
    {
        public string Attorney { get; set; }
        public int id { get; set; }
        public string lname { get; set; }
        public string fname { get; set; }

        public int procedureDetail_id { get; set; }
        public string gender { get; set; }
        public string name { get; set; }
        public DateTime? dob { get; set; }
        public DateTime? doe { get; set; }
        public DateTime? doa { get; set; }
        public string mc { get; set; }
        public string mc_details { get; set; }
        public string acct_no { get; set; }
        public string phone { get; set; }
        public string PolicyNo { get; set; }
        public string primary_claim_no { get; set; }
        public string cmpname { get; set; }
        public string allergies { get; set; }
        public string note { get; set; }
        
        public string casetype { get; set; }
        public string location { get; set; }
        public string vaccinated { get; set; }
        public string mcode { get; set; }
        public string sides { get; set; }
        public string level { get; set; }
        public string providerName { get; set; }
        public int? surgercy_center { get; set; }
        public string surgon_name { get; set; }
        public string assistent_name { get; set; }
        public string Surgerycenter_name { get; set; }
        public DateTime? requested { get; set; }
        public DateTime? executed { get; set; }
        public DateTime? scheduled { get; set; }
        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }
        public string sx_center_name { get; set; }
        public List<POCReportVM> lstPOCReport { get; set; }

        public int? locationid { get; set; } 
        public int? mcodeid { get; set; }

        public bool _requested { get; set; }
        public bool _executed { get; set; }
        public bool _scheduled { get; set; }
       
    }
}
