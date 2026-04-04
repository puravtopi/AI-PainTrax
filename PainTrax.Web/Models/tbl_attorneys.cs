using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_attorneys
	{
		[Required(ErrorMessage = "Please enter Attorney EmailId")]
		public string? EmailId { get; set; }
		[Required(ErrorMessage = "Please enter Attorney Name")]
		public string Attorney { get; set; }		
		public string? ContactNo { get; set; }
		public string? ContactName { get; set; }
		public int? Id { get; set; }
		public string? Address { get; set; }
		public string? City { get; set; }
		
		public string? State { get; set; }
		public string? Zip { get; set; }
		public string? Paralegal { get; set; }
		public DateTime? CreatedDate { get; set; }
		public int? CreatedBy { get; set; }
		public int? old_id { get; set; }
		public int? cmp_id { get; set; }

        public string? label { get; set; }
        public string? val { get; set; }

    }
}