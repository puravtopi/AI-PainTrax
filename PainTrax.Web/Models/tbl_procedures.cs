using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
	public class tbl_procedures
	{
		public string? position { get; set; }
		public string? mcode { get; set; }	
		public string? heading { get; set; }
		public int? display_order { get; set; }				

		[Required(ErrorMessage = "Please enter Body Part")]
		public string bodypart { get; set; }
		public int? inhouseproc { get; set; }
		public int? id { get; set; }
		public string? ccdesc { get; set; }
		public string? pedesc { get; set; }
		public string? adesc { get; set; }
		public string? pdesc { get; set; }
		public bool cf { get; set; }
		public bool pn { get; set; }
		public bool? preselect { get; set; }		
		
		public bool inhouseprocbit { get; set; }
		public bool haslevel { get; set; }
		public string? hasmuscle { get; set; }
		public string? hasmedication { get; set; }
		public int? bid { get; set; }
		public bool inout { get; set; }
		public bool sides { get; set; }
		public bool? status { get; set; }
		public string? hassubprocedure { get; set; }
		public string? s_ccdesc { get; set; }
		public string? s_pedesc { get; set; }
		public string? s_adesc { get; set; }
		public string? s_pdesc { get; set; }
		public string? e_ccdesc { get; set; }
		public string? e_pedesc { get; set; }
		public string? e_adesc { get; set; }
		public string? e_pdesc { get; set; }
		public string? s_heading { get; set; }
		public string? e_heading { get; set; }
		public string? levelsdefault { get; set; }
		public string? sidesdefault { get; set; }
		public string? mcode_desc { get; set; }
		public int? cmp_id { get; set; }
		public int? createdby { get; set; }
		public DateTime? createddate { get; set; }
		public string? upload_template { get; set; }
        public string? injection_description { get; set; }

    }
}