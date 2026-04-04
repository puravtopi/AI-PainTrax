using PainTrax.Web.Helper;
using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{

    public class tbl_pocconfig
    {
        public string? id { get; set; }
        public string? columns { get; set; }
        public int? cmp_id { get; set; }
        public string? export_Columns { get; set; }
        public List<CheckBoxItem> Listcolumns { get; set; }

    }
}