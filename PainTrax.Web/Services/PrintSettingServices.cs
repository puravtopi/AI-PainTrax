using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class PrintSettingServices : ParentService
    {
        public PrintSettingServices()
        {
        }

        public List<tbl_print_label> GetAll(string cnd = "")
        {
            string query = "select * from tbl_print_label where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            List<tbl_print_label> dataList = ConvertDataTable<tbl_print_label>(GetData(query));
            return dataList;
        }

        public List<tbl_print_label> GetAautoComplete(string cnd = "")
        {
            string query = "select *,attorney as label,attorney as val from tbl_print_label where 1=1 ";
            query += cnd;
            List<tbl_print_label> dataList = ConvertDataTable<tbl_print_label>(GetData(query));
            return dataList;
        }

        public tbl_print_label? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_print_label where Id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", id);
            var datalist = ConvertDataTable<tbl_print_label>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public string GetCodeTitle(string code, string cmp_id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_print_label where cmp_id=@cmp_id and lbl_code=@code", conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);
            cm.Parameters.AddWithValue("@code", code);
            var datalist = ConvertDataTable<tbl_print_label>(GetData(cm)).FirstOrDefault();

            string title = "";

            if (datalist.is_show.Value)
            {
                if (datalist.style.ToLower() == "bold")
                    title = "<b>" + datalist.lbl_title + "</b> ";
                else if (datalist.style.ToLower() == "italic")
                    title = "<i>" + datalist.lbl_title + "</i> ";
                else
                    title = "<b><i>" + datalist.lbl_title + "</i></b> ";

                if (datalist.is_new_line.Value)
                    title = title + "<br/> ";

            }


            return title;


        }

        public void Update(tbl_print_label data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_print_label SET
		lbl_title=@lbl_title,style=@style,is_new_line=@is_new_line,
		is_show=@is_show 
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            cm.Parameters.AddWithValue("@lbl_title", data.lbl_title);
            cm.Parameters.AddWithValue("@is_show", data.is_show);
            cm.Parameters.AddWithValue("@style", data.style);
            cm.Parameters.AddWithValue("@is_new_line", data.is_new_line);

            Execute(cm);
        }

        public bool ShowTitle(string code, string cmp_id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_print_label where cmp_id=@cmp_id and lbl_code=@code", conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);
            cm.Parameters.AddWithValue("@code", code);
            var datalist = ConvertDataTable<tbl_print_label>(GetData(cm)).FirstOrDefault();

            if (datalist.is_show.Value)
                return true;
            else
                return false;
        }

        public string GetTitle(string code, string cmp_id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_print_label where cmp_id=@cmp_id and lbl_code=@code", conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);
            cm.Parameters.AddWithValue("@code", code);
            var datalist = ConvertDataTable<tbl_print_label>(GetData(cm)).FirstOrDefault();

            return datalist.lbl_title;
        }

        public tbl_template? GetTemplate(string cmp_id,string type)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_template where cmp_id=@cmp_id and type=@type", conn);
            cm.Parameters.AddWithValue("@cmp_id", cmp_id);
            cm.Parameters.AddWithValue("@type", type);

            var datalist = ConvertDataTable<tbl_template>(GetData(cm)).FirstOrDefault();

            return datalist;

        }

        public tbl_template? GetOneTemplate(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_template where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
         

            var datalist = ConvertDataTable<tbl_template>(GetData(cm)).FirstOrDefault();

            return datalist;

        }

        public List<tbl_template> GetAllTemplate(string cnd = "")
        {
            string query = "select * from tbl_template where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            MySqlCommand cm = new MySqlCommand(query, conn);
           

            var datalist = ConvertDataTable<tbl_template>(GetData(cm));

            return datalist;

        }

        public void SaveTemplate(tbl_template model)
        {

            if (model.id > 0)
            {
                MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_template SET
		content=@content
		where id=@id", conn);
                cm.Parameters.AddWithValue("@id", model.id);

                cm.Parameters.AddWithValue("@content", model.content);

                Execute(cm);
            }
            else {
                MySqlCommand cm = new MySqlCommand(@"insert into tbl_template(type,content,cmp_id)Values(@type,@content,@cmp_id);select 1;", conn);
                cm.Parameters.AddWithValue("@type", model.type);

                cm.Parameters.AddWithValue("@content", model.content);
                cm.Parameters.AddWithValue("@cmp_id", model.cmp_id);

                Execute(cm);
            }
        }
    }
}
