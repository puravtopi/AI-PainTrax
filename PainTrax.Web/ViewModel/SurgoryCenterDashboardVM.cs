namespace PainTrax.Web.ViewModel
{
    public class SurgoryCenterDashboardVM
    {
        public string? ScheduledDate { get; set; }
        public string? SurgerycenterName { get; set; }
        public string? ProcedureDetailIDs { get; set; }
        public string? SurgonName { get; set; }
        public int? SurgerycenterId { get; set; }
        public long? ScheduledCnt { get; set; }
        public long? BookedCnt { get; set; }
        public long? ExecutedCnt { get; set; }
        public DateTime? fdate { get; set; }
        public DateTime? tdate { get; set; }
        public List<SurgoryCenterDashboardVM> lstSurgoryCenterDashboardVM { get; set; } = new();
    }
}
