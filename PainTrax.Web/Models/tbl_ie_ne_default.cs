namespace MS.Models
{    
    public class tbl_ie_ne_default
    {
        public int id { get; set; }
        public string? neurological_exam { get; set; }
        public string? sensory { get; set; }
        public string? manual_muscle_strength_testing { get; set; }
        public string? other_content { get; set; }
       
        public int? cmp_id { get; set; }
    }
}
