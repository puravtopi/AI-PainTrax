using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{

    public class tbl_macros_master
    {
       
        [Required(ErrorMessage ="Please Enter BodyPart")]
        public string? bodypart { get; set; }
        public string? heading { get; set; }
        public int? id { get; set; }
        public string? cc_desc { get; set; }
        public string? pe_desc { get; set; }
        public string? a_desc { get; set; }
        public string? rom_desc { get; set; }
        public string? p_desc { get; set; }
        public bool? cf { get; set; }
        public bool? pn { get; set; }
        public bool? pre_select { get; set; }        
        public int? cmp_id { get; set; }
        public DateTime? created_date { get; set; }
        public int? created_by { get; set; }
        public DateTime? updated_date { get; set; }
        public int? updated_by { get; set; }

        // Extra fields for Pre Op

        public string? pc_desc { get; set; }
        public string? ros_desc { get; set; }
        public string? ds_desc { get; set; }
        public string? pt_desc { get; set; }
        public string? drd_desc { get; set; }
        public string? drd_notes { get; set; }

    }
}