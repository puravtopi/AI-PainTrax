using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_locations
	{
		public string nameofpractice { get; set; }

		[Required(ErrorMessage = "Please Enter Location")]
		public string location { get; set; }
		public int? id { get; set; }
		public string? emailid { get; set; }
		public string? telephone { get; set; }
		public string? contactpersonname { get; set; }
		public bool? setasdefault { get; set; }
		public string? address { get; set; }
		public string? city { get; set; }
		public string? state { get; set; }
		public string? zipcode { get; set; }		
		
		public string? fax { get; set; }
		public string? drfname { get; set; }
		public string? drlname { get; set; }
		public bool isactive { get; set; }
		public int? cmp_id { get; set; }
		public DateTime? createddate { get; set; }
		public int? createdby { get; set; }
		public DateTime? updatedate { get; set; }
		public int? updatedby { get; set; }
		public int? old_id { get; set; }

		public string? header_template { get; set; }
		public string? header_template_hidden { get; set; }
	}
}