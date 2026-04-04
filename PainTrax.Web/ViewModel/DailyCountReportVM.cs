using Microsoft.VisualBasic;

namespace PainTrax.Web.ViewModel
{
    public class DailyCountReportVM
    {
        public DateTime? doe { get; set; }

        public string location { get; set; }
        public Int64 WC { get; set; }
        public Int64 NF { get; set; }
        public Int64 LIEN { get; set; }
        public Int64 NoOFIE { get; set; }


        public Int64 NoOFFU { get; set; }

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<DailyCountReportVM> lstDailyCountReport { get; set; }


    }
}
