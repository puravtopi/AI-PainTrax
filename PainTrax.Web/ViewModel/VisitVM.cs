using MS.Models;
using PainTrax.Web.Models;
using System.ComponentModel.DataAnnotations;
using static PainTrax.Web.Helper.EnumHelper;

namespace PainTrax.Web.ViewModel
{
    public class VisitVM
    {
        public int? id { get; set; }
        public int? fu_id { get; set; }
        public int? patientid { get; set; }
        public int? locationid { get; set; }
        public int? providerid { get; set; } 
        public int? physicianid {  get; set; }
        public string? providerName { get; set; }

        //[DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime? dos { get; set; }
        public DateTime? old_dos { get; set; }

        public DateTime? doi { get; set; }

        //[DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime? dov { get; set; }

        //[DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime? doa { get; set; }
        public string? compensation { get; set; }
        public string? accidentType { get; set; }
        public string? fname { get; set; }
        public string? lname { get; set; }
        public string? mname { get; set; }
        public string? gender { get; set; }
        public Gender gendername { get; set; }

        //[DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime? dob { get; set; }
        public int? age { get; set; }
        public string? email { get; set; }
        public string? handeness { get; set; }
        public string? ssn { get; set; }
        public string? address { get; set; }
        public string? city { get; set; }
        public string? state { get; set; }
        public string? zip { get; set; }
        public string? home_ph { get; set; }
        public string? mobile { get; set; }
        public string? history { get; set; }
        public bool? vaccinated { get; set; }
        public string? mc { get; set; }
        public string? mc_details { get; set; }
        public string? account_no { get; set; }
        public string? prime_cmpname { get; set; }
        public string? prime_address { get; set; }
        public string? prime_phone { get; set; }
        public string? prime_claim_no { get; set; }
        public string? prime_policy_no { get; set; }
        public string? prime_WCB_group { get; set; }

        public string? sec_cmpname { get; set; }
        public string? sec_address { get; set; }
        public string? sec_phone { get; set; }
        public string? sec_claim_no { get; set; }
        public string? sec_policy_no { get; set; }
        public string? sec_WCB_group { get; set; }

        public string? emp_name { get; set; }
        public string? emp_address { get; set; }
        public string? emp_phone { get; set; }
        public string? emp_fax_no { get; set; }

        public string? attory_name { get; set; }
        public string? attory_phone { get; set; }
        public string? attory_fax_no { get; set; }
        public string? attory_email { get; set; }
        public string? paralegal { get; set; }

        public string? adj_name { get; set; }
        public string? adj_phone { get; set; }
        public string? adj_email { get; set; }
        public string? adj_fax_no { get; set; }

        public string? alert_note { get; set; }
        public string? location { get; set; }
        public string? casetype { get; set; }
        public string? referring_physician { get; set; }
        public string? type { get; set; }
        public string? doc_json { get; set; }
        public string? procedure_performed { get; set; }
        public tbl_pre pre { get; set; }
        public tbl_post post { get; set; }
        public tbl_ie_page1 Page1 { get; set; }
        public tbl_ie_page2 Page2 { get; set; }
        public tbl_ie_page3 Page3 { get; set; }
        public tbl_ie_ne NE { get; set; }
        public tbl_ie_other Other { get; set; }
        public tbl_ie_comment Comment { get; set; }
        public List<tbl_treatment_master> listTreatments { get; set; }
        public List<WebisteMarcoVM> listWebsiteMacros { get; set; }

    }
}
