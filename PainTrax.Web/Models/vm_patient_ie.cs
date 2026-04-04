namespace PainTrax.Web.Models;

public class vm_patient_ie
{
    internal string? casetype;
   
    public string? fname { get; set; }
    public string? lname { get; set; }
    public DateTime? dob { get; set; }
    public DateTime? doa { get; set; }
    public DateTime? doe { get; set; }
    public string? compensation { get; set; }
    public string? accidentType { get; set; }
    public string? location { get; set; }
    public int? id { get; set; }
    public int? patient_id { get; set; }
    public int? location_id { get; set; }
    public int? attorney_id { get; set; }
    public int? primary_ins_cmp_id { get; set; }
    public int? secondary_ins_cmp_id { get; set; }
    public int? emp_id { get; set; }
    public int? adjuster_id { get; set; }
    public int? provider_id { get; set; }
    public string? primary_claim_no { get; set; }
    public string? secondary_claim_no { get; set; }
    public string? providerName { get; set; }
    public string? primary_policy_no { get; set; }
    public string? secondary_policy_no { get; set; }

    public string? primary_wcb_group { get; set; }
    public string? secondary_wcb_group { get; set; }
    public string? note { get; set; }
    public string? ins_note { get; set; }
    public string? alert_note { get; set; }
    public DateTime? created_date { get; set; }
    public int? created_by { get; set; }
    public DateTime? updated_date { get; set; }

    public int? physicianid { get; set; }
    public int? updated_by { get; set; }
    public bool? is_active { get; set; }
    public int? patientId { get; set; }

    public string? mname { get; set; }
    public string? gender { get; set; }

    public string? email { get; set; }
    public string? handeness { get; set; }
    public string? ssn { get; set; }
    public string? address { get; set; }
    public string? city { get; set; }
    public string? state { get; set; }
    public string? zip { get; set; }
    public string? home_ph { get; set; }
    public string? mobile { get; set; }
    public string? account_no { get; set; }
    public string? strlocation { get; set; }
    public bool? vaccinated { get; set; }


    public bool? mc { get; set; }
    public bool? isFU { get; set; }
    public int? age { get; set; }
    public int? intakeid { get; set; }
    public string? procedure_performed { get; set; }

}