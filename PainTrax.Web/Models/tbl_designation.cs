using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_designation
	{
        //[Required(ErrorMessage = "Please enter Designation code")]
        public string? code { get; set; }

		[Required(ErrorMessage = "Please enter Designation Name")]
		public string? title { get; set; }
		public int? id { get; set; }
		public int? cmp_id { get; set; }
		public DateTime? createddate { get; set; }
		public int? createdby { get; set; }
		public DateTime? updateddate { get; set; }
		public int? updatedby { get; set; }
		public int? old_id { get; set; }
		public int? injection_description { get; set; }

	}
}