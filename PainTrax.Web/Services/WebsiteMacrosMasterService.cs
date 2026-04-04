using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;
using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Services
{
    public class WebsiteMacrosMasterService : ParentService
    {

        public List<tbl_website_marco> GetAll(string cnd = "")
        {
            string query = "select * from tbl_website_marco t where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            List<tbl_website_marco> dataList = ConvertDataTable<tbl_website_marco>(GetData(query));
            return dataList;
        }
        public List<tbl_website_marco> GetAautoComplete(string cnd = "")
        {

            string query = "select *,t.key as label,t.desc as val from tbl_website_marco t where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;

            List<tbl_website_marco> dataList = ConvertDataTable<tbl_website_marco>(GetData(query));
            return dataList;
        }

        public List<WebisteMarcoVM> GetAutoComplete(int cmpid, string type = "All")
        {
            string query = "select * from tbl_website_marco t where cmp_id=" + cmpid + " and type='" + type + "' order by t.key";

            List<tbl_website_marco> data = ConvertDataTable<tbl_website_marco>(GetData(query));

            List<WebisteMarcoVM> dataList = new List<WebisteMarcoVM>();

            foreach (var item in data)
            {
                var obj = new WebisteMarcoVM()
                {
                    id = "~" + item.desc,
                    userId = item.id.ToString(),
                    name = item.key
                };

                dataList.Add(obj);
            }

            return dataList;
        }


        public tbl_website_marco? GetOne(tbl_website_marco data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_website_marco where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_website_marco>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public void Insert(tbl_website_marco data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_website_marco
		(`key`,`type`,`desc`,`cmp_id`,`user_id`)Values
				(@key,@type,@desc,@cmp_id,@user_id)", conn);
            cm.Parameters.AddWithValue("@key", data.key);
            cm.Parameters.AddWithValue("@desc", data.desc);
            cm.Parameters.AddWithValue("@type", data.type);
            cm.Parameters.AddWithValue("@user_id", 0);

            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);


            Execute(cm);
        }
        public void Update(tbl_website_marco data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_website_marco SET
		`key`=@key,
        `type`=@type,
        `user_id`=@user_id,
		`desc`=@desc where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@key", data.key);
            cm.Parameters.AddWithValue("@desc", data.desc);
            cm.Parameters.AddWithValue("@type", data.type);
            cm.Parameters.AddWithValue("@user_id", 0);
            Execute(cm);
        }
        public void Delete(tbl_website_marco data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_website_marco
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }
    }
}
