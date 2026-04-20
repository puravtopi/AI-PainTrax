using Newtonsoft.Json;          // or System.Text.Json — swap as needed
using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.ViewModel
{
    public class PatientImportReportVM
    {
        [Key]
        public int id { get; set; }

        [Required]
        [MaxLength(100)]
        public string last_name { get; set; } = string.Empty;

        [MaxLength(100)]
        public string? first_name { get; set; }

        public DateTime? dob { get; set; }

        [MaxLength(10)]
        public string? sex { get; set; }

        [MaxLength(255)]
        public string? address { get; set; }

        [MaxLength(20)]
        public string? phone { get; set; }

        [MaxLength(15)]
        public string? ssn { get; set; }

        [MaxLength(150)]
        public string? employer_company { get; set; }

        [MaxLength(255)]
        public string? employer_address { get; set; }

        [MaxLength(100)]
        public string? emergency_contact { get; set; }

        [MaxLength(20)]
        public string? work_phone { get; set; }

        //public DateTime? date_of_accident { get; set; }
        public DateTime? doa { get; set; }

        [MaxLength(100)]
        public string? condition_related_to { get; set; }

        [MaxLength(150)]
        public string? insurance_company { get; set; }

        [MaxLength(255)]
        public string? insurance_address { get; set; }

        [MaxLength(20)]
        public string? insurance_phone { get; set; }

        [MaxLength(50)]
        public string? claim_number { get; set; }

        [MaxLength(255)]
        public string? claim_address { get; set; }

        [MaxLength(5)]
        public string? nf2 { get; set; }

        [MaxLength(50)]
        public string? policy_number { get; set; }

        [MaxLength(100)]
        public string? policy_holder { get; set; }

        [MaxLength(50)]
        public string? wcb_number { get; set; }

        [MaxLength(100)]
        public string? carrier_case_number { get; set; }

        [MaxLength(150)]
        public string? policy_adjuster { get; set; }

        [MaxLength(150)]
        public string? attorney { get; set; }

        [MaxLength(150)]
        public string? firm_name { get; set; }

        [MaxLength(255)]
        public string? attorney_address { get; set; }

        [MaxLength(20)]
        public string? attorney_phone { get; set; }

        [MaxLength(20)]
        public string? attorney_fax { get; set; }

        public DateTime imported_at { get; set; } = DateTime.Now;

        public int? cmpy_id { get; set; }

        public int? loc_id { get; set; }

        //extra columns added as per charse and arun.  
        public DateTime? DOE { get; set; }
 

        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }

        public List<PatientImportReportVM> lstPatientImportReport { get; set; }

        public int? locationid { get; set; }


    }

    // ============================================================
    //  PatientImportRowVM  –  matches the Excel column names that
    //  SheetJS sends as JSON keys (must match COLUMNS array in the
    //  view, mapped from "Patient Name" → PatientName etc.)
    // ============================================================
    public class PatientImportRowVM
    {
        [JsonProperty("Patient Name")] public string PatientName { get; set; }
        [JsonProperty("DOB")] public string DOB { get; set; }
        [JsonProperty("Sex")] public string Sex { get; set; }
        [JsonProperty("Address")] public string Address { get; set; }
        [JsonProperty("Phone")] public string Phone { get; set; }
        [JsonProperty("Social Security #")] public string SocialSecurityNo { get; set; }
        [JsonProperty("Employer/Company")] public string EmployerCompany { get; set; }
        [JsonProperty("Employer Address")] public string EmployerAddress { get; set; }
        [JsonProperty("Emergency Name")] public string EmergencyName { get; set; }
        [JsonProperty("Work Phone")] public string WorkPhone { get; set; }
        [JsonProperty("Date of Accident")] public string DateOfAccident { get; set; }
        [JsonProperty("CaseType")] public string CaseType { get; set; }
        [JsonProperty("Insurance Company")] public string InsuranceCompany { get; set; }
        [JsonProperty("Ins Address")] public string InsAddress { get; set; }
        [JsonProperty("Ins Phone")] public string InsPhone { get; set; }
        [JsonProperty("Claim #")] public string ClaimNo { get; set; }
        [JsonProperty("Claim Address")] public string ClaimAddress { get; set; }
        [JsonProperty("NF-2")] public string NF2 { get; set; }
        [JsonProperty("Policy #")] public string PolicyNo { get; set; }
        [JsonProperty("Policy Holder")] public string PolicyHolder { get; set; }
        [JsonProperty("WCB #")] public string WCBNo { get; set; }
        [JsonProperty("Carrier Case #")] public string CarrierCaseNo { get; set; }
        [JsonProperty("Policy Adjuster")] public string PolicyAdjuster { get; set; }
        [JsonProperty("Attorney")] public string Attorney { get; set; }
        [JsonProperty("Firm Name")] public string FirmName { get; set; }
        [JsonProperty("Attorney Address")] public string AttorneyAddress { get; set; }
        [JsonProperty("Attorney Phone")] public string AttorneyPhone { get; set; }
        [JsonProperty("Attorney Fax")] public string AttorneyFax { get; set; }

        [JsonProperty("DOE")] public string DOE { get; set; }
        
    }
}
