namespace MS.Models;

public class tbl_patient_fu{

	public int? id {get;set;} 
	public int? patient_id {get;set;}
	public int? provider_id { get; set; }
	public int? patientIE_ID {get;set;}
	public int? location_id {get;set;}
	public int? physicianid { get;set;}

    public DateTime? doe {get;set;}

	public DateTime? created_date {get;set;}
	public int? created_by {get;set;}
	public DateTime? updated_date {get;set;}
	public int? updated_by {get;set;}
	public bool? is_active {get;set;}
	public int? cmp_id {get;set;}
	public string? extra_comments {get;set;}
	public string? type { get;set;}
	public string? history { get;set;}
    public string? accident_type { get; set; }
    public bool? final_save { get; set; }

    public string? procedure_performed { get; set; }



}