namespace PainTrax.Web.ViewModel
{
	public class AppointmentsVM
	{

		public int? app_id { get; set; }
		public int? provider_id { get; set; }
		public int? patient_id { get; set; }
		public int? location_id { get; set; }
		public string? app_date { get; set; }
		public string? app_time { get; set; }
		public string? app_note { get; set; }
		public int? status_id { get; set; }
		public string? tags { get; set; }
        public string? app_fromdate { get; set; }
        public string? app_todate { get; set; }
        public string? app_multitime { get; set; }
        public string? app_days { get; set; }

        public int? cmp_id { get; set; }

		public int? isEdit { get;set; }

    }
}
