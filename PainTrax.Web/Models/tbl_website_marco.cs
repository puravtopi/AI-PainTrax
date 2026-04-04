using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
    public class tbl_website_marco
    {
        public string desc { get; set; }
        public string key { get; set; }
        [Required(ErrorMessage = "Please Enter Type")]
        public string type { get; set; }
       
        public int id { get; set; }
        public string label { get; set; }
        public string val { get; set; }
        public int? cmp_id { get; set; }
        //public int? user_id { get; set; }
    }
}
