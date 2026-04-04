using Org.BouncyCastle.Asn1.Ocsp;
using System.ComponentModel.DataAnnotations;
using static PainTrax.Web.Helper.EnumHelper;

namespace PainTrax.Web.Models
{
    public class SI_Report
    {
        public Int32 ieid { get; set; }
        public Int64 fuid { get; set; }

        public string? name { get; set; }
        public string? fname { get; set; }
        public string? lname { get; set; }
        public string?  visitiefu { get; set; }
        public string? followupdate { get; set; }
        public DateTime? doa { get; set; }
        public DateTime ? doe { get; set; }
        public string? casetype { get; set; }
        public string? location { get; set; }
        public string? requested { get; set; }
        public string? scheduled { get; set; }
        public string? alert { get; set; }
        public string? inhouse { get; set; }
        public string? insuranceCompany { get; set; }


        /*public string label { get; set; }
        public string val { get; set; }*/

    }
}