namespace MS.Models;

public class tbl_fu_other{

	public int? id {get;set;}
	public int? ie_id {get;set;}
	public int? fu_id {get;set;}
	public int? patient_id {get;set;}
	public string? treatment_details {get;set;}
    public string? note1 { get; set; }
    public string? note2 { get; set; }
    public string? note3 { get; set; }
    public string? treatment_delimit {get;set;}

    public string? treatment_delimit_desc { get; set; }

    public string? followup_duration {get;set;}
	public DateTime? followup_date {get;set;}

}