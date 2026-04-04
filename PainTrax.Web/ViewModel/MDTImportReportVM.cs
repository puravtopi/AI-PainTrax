using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class MDTImportReportVM
    {
        public DateTime? doe { get; set; }
        public Int32? PatientIE_ID { get; set; }
        public string lname { get; set; }
        public string fname { get; set; }
        public string mname { get; set; }
        public string gender { get; set; }
        public DateTime? dob { get; set; }
        public string address { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string zip { get; set; }
        public string home_ph { get; set; }
        public string mobile { get; set; }
        public string location { get; set; }
        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<MDTImportReportVM> lstMDTImportReport { get; set; }


    }
}
