using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_diagcodes
	{
		public string? BodyPart { get; set; }		
		public string? DiagCode { get; set; }
		public string? Description { get; set; }
		public int? display_order { get; set; }
		public int? Id { get; set; }
		public DateTime? CreatedDate { get; set; }
		public int? CreatedBy { get; set; }
		public bool PreSelect { get; set; }
		
		public int? old_id { get; set; }
		public int? cmp_id { get; set; }

	}
}