namespace PainTrax.Web.Models
{
    public class tbl_settings
    {
        public int id { get; set; }
        public int page_size { get; set; }
        public int? location { get; set; }        
        public int cmp_id { get; set; }
        public string? dateformat { get; set; }
        public bool? pageBreakForInjection { get; set; }
        public bool? injectionAsSeparateBlock { get; set; }
        public bool isdaignosisshow { get; set; }
        public string? foundStatment { get; set; }
        public string? notfoundStatment { get; set; }
        public string? header_template { get; set; }
        public string? header_template_hidden { get; set; }
        public string? font_family { get; set; }
        public string? font_size { get; set; }        
        public string? gait_default { get; set; }        
        public string? fu_default { get; set; }        
        public bool show_preop { get; set; }        
        public bool show_postop { get; set; }
        public bool table_border { get; set; }
        public string? sign_content { get; set; }
        public string? casetype { get; set; }
        public string? diagcervialbulge_comma { get; set; }
        public string? diagthoracicbulge_comma { get; set; }
        public string? diaglumberbulge_comma { get; set; }
        public string? diagleftshoulder_comma { get; set; }
        public string? diagrightshoulder_comma { get; set; }
        public string? diagleftknee_comma { get; set; }
        public string? diagrightknee_comma { get; set; }
        public string? diagbrain_comma { get; set; }
        public string? other1_comma { get; set; }
        public string? other2_comma { get; set; }
        public string? other3_comma { get; set; }
        public string? other4_comma { get; set; }
        public string? other5_comma { get; set; }
        public string? other6_comma { get; set; }
        public string? other7_comma { get; set; }
       
    }
}
