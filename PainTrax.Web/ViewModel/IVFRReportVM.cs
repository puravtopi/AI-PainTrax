using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class IVFRReportVM
    {
        public int procedureDetail_id { get; set; }
        public string gender { get; set; }
        public string name { get; set; }
        public DateTime? dob { get; set; }
        public DateTime? doa { get; set; }
        public string mc { get; set; }
        public string phone { get; set; }
        public string primary_policy_no { get; set; }
        public string primary_claim_no { get; set; }
        public string cmpname { get; set; }
        
        public string casetype { get; set; }
        public string location { get; set; }
        public string vaccinated { get; set; }
        public string mcode { get; set; }
        public string ssn { get; set; }
        public string Address { get; set; }
        public string InsCo { get; set; }


        public DateTime? requested { get; set; }
        public DateTime? executed { get; set; }
        public DateTime? scheduled { get; set; }
        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<IVFRReportVM> lstIVFRReport { get; set; }

        public int? locationid { get; set; } 
        public int? mcodeid { get; set; }

        //public bool _requested { get; set; }
        //public bool _executed { get; set; }
        public bool _scheduled { get; set; }
    }
}
