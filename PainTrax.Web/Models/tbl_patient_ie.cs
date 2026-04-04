namespace MS.Models;

public class tbl_patient_ie{

	public int? id {get;set;}
	public int? patient_id {get;set;}
	public int? location_id {get;set;}
	public int? attorney_id {get;set;}
	public int? primary_ins_cmp_id {get;set;}
	public int? secondary_ins_cmp_id {get;set;}
	public int? emp_id {get;set;}
	public int? adjuster_id {get;set;}
    public int? provider_id { get; set; }
    public DateTime? doe {get;set;}
	public DateTime? doa {get;set;}
	public string? primary_claim_no {get;set;}
	public string? secondary_claim_no {get;set;}
	public string? primary_policy_no {get;set;}
	public string? secondary_policy_no {get;set;}
	public string? compensation {get;set;}
	public string? primary_wcb_group { get;set;}
	public string? secondary_wcb_group { get;set;}
	public string? note {get;set;}
	public string? poc_note {get;set;}
	public string? ins_note {get;set;}
	public string? alert_note {get;set;}
	public string? referring_physician { get;set;}
	public string? accident_type { get;set;}
	
	public DateTime? created_date {get;set;}
	public int? created_by {get;set;}
	public int? updated_by {get;set;}
	public DateTime? updated_date {get;set;}
	public bool? is_active {get;set; }
	public string? state { get; set; }
	public int? physicianid { get;set;}
	public int? intakeid { get;set;}

    public string? procedure_performed { get; set; }

}