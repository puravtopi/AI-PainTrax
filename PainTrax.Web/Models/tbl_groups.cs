using PainTrax.Web.Helper;
using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_groups
	{
        public string? pages_name { get; set; }

        [Required(ErrorMessage = "Please Enter GroupTitle")]
        public string? Title { get; set; }

        [Required(ErrorMessage = "Please Select Location")]
        public string? location_name { get; set; }
       
        public string? report_name { get; set; }
        public string? Location_ids { get; set; }
        public string? pages_ids { get; set; }
        public string? reports_ids { get; set; }
        public string? role_ids { get; set; }
        public string? role_name { get; set; }
        public string? form_name { get; set; }
        public DateTime? CreatedDate { get; set; }
		public int? CreatedBy { get; set; }
		public int? old_id { get; set; }
		public int? cmp_id { get; set; }
        public int? Id { get; set; }
        public List<CheckBoxItem> LocationList { get; set; }
		public List<CheckBoxItem> PagesList { get; set; }
		public List<CheckBoxItem> ReportsList { get; set; }
		public List<CheckBoxItem> RoleList { get; set; }
		public List<CheckBoxItem> FormList { get; set; }

	}
}