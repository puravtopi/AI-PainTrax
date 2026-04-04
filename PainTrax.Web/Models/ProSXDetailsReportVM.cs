using Microsoft.VisualBasic;
using System.Security.Cryptography.X509Certificates;

namespace PainTrax.Web.ViewModel
{
    public class ProSXDetailsReportVM
    {
        public string Attorney { get; set; }
        public DateTime? DOE { get; set; }
        public string PolicyNo { get; set; }
        public int ProcedureDetail_ID { get; set; }
        public string sex { get; set; }
        public string name { get; set; }
        public string mc { get; set; }
        public string casetype { get; set; }
        public string location { get; set; }
        public string vaccinated { get; set; }

        public string mcode { get; set; }
        public string bodypart { get; set; }
        public string ins_ver_status { get; set; }
        public string mc_status { get; set; }
        public string case_status { get; set; }
        public string insverstatus { get; set; }
        public string vac_status { get; set; }

        //public DateTime? requested { get; set; }
        //public DateTime? executed { get; set; }
        //public DateTime? scheduled { get; set; }
        public string? requested { get; set; }
        public string? executed { get; set; }
        public string? scheduled { get; set; }

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<ProSXDetailsReportVM> lstProSXDetailsReport { get; set; }

        public int? locationid { get; set; } 
       
    }
}
