namespace PainTrax.Web.ViewModel
{
    public class DaignoCodeVM
    {
        public int DaignoCodeId { get; set; }
        public bool IsSelect { get; set; }
        public string DiagCode { get; set; }
        public string? Description { get; set; }

        public int? Display_Order { get; set; }
        public int? cmp_id { get; set; }
    }
}
