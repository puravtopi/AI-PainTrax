namespace PainTrax.Web.ViewModel
{
    public class IEToFuVM
    {
        public string? tbl_name { get; set; }

        public string? tbl_column { get; set;}

        public int? created_by { get; set; }
        public int cmp_id { get; set; }
    }

    public class IEToFuVMList
    {
        public List<IEToFuVM> SOAP { get; set; }
        public List<IEToFuVM> page2 { get; set; }
        public List<IEToFuVM> page3 { get; set; }
        public List<IEToFuVM> ne { get; set; }
        public List<IEToFuVM> other { get; set; }
        public List<IEToFuVM> dg { get; set; }

        public int cntSOAP { get; set; }
        public int cntPage2 { get; set; }
        public int cntPageNE { get; set; }
        public int cntPageOther { get; set; }
        public int cntPage3 { get; set; }
        public int cntPageDaignoList { get; set; }
    }
}
