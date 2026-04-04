namespace PainTrax.Web.Models;

public class tbl_patient{

	public int? id {get;set;}
	public string? fname {get;set;}
	public string? lname {get;set;}
	public string? mname {get;set;}
	public string? gender {get;set;}
	public DateTime? dob {get;set;}
	public int? age {get;set;}
	public string? email {get;set;}
	public string? handeness {get;set;}
	public string? ssn {get;set;}
	public string? address {get;set;}
	public string? city {get;set;}
	public string? state {get;set;}
	public string? zip {get;set;}
	public string? home_ph {get;set;}
	public string? mobile {get;set;}
	public bool? vaccinated {get;set;}
	public string? mc {get;set;}
	public string? account_no {get;set;}
	public DateTime? createddate {get;set;}
	public int? createdby {get;set;}
	public DateTime? updatedate {get;set;}
	public int? updatedby {get;set;}
	public int? cmp_id {get;set;}
	public string? label { get; set; }
	public string? val { get; set; }
	public int? physicianid { get;set;}

    public string? mc_details { get; set; }
}