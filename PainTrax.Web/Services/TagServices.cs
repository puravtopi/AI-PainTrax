using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class TagServices : ParentService
    {
        public List<tbl_tag> GetAll(string cnd = "")
        {
            string query = "select * from tbl_tags where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;

            List<tbl_tag> dataList = ConvertDataTable<tbl_tag>(GetData(query));
            return dataList;
        }
    }
}
