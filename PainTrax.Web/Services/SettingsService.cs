using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class SettingsService : ParentService
    {
        public List<tbl_settings> GetAll(string cnd = "")
        {

            string query = "select * from tbl_settings where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            List<tbl_settings> dataList = ConvertDataTable<tbl_settings>(GetData(query));
            return dataList;
        }
        public tbl_settings? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_settings where cmp_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_settings>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
        public void Insert(tbl_settings data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_settings
		(page_size,location,dateformat,cmp_id,pageBreakForInjection,injectionAsSeparateBlock,font_family,font_size)Values
				(@page_size,@location,@dateformat,@cmp_id,@pageBreakForInjection,@injectionAsSeparateBlock,@isdaignosisshow,
            @foundStatment,@notfoundStatment,@header_template,@font_family,@font_size,
            @diagcervialbulge_comma,@diagthoracicbulge_comma,@diaglumberbulge_comma,
            @diagleftshoulder_comma,@diagrightshoulder_comma,@diagleftknee_comma,
            @diagrightknee_comma,@diagbrain_comma,@other1_comma,
            @other2_comma,@other3_comma,@other4_comma,
            @other5_comma,@other6_comma,@other7_comma)", conn);
            cm.Parameters.AddWithValue("@page_size", data.page_size);
            cm.Parameters.AddWithValue("@location", data.location);
            cm.Parameters.AddWithValue("@dateformat",data.dateformat);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@pageBreakForInjection", data.pageBreakForInjection);
            cm.Parameters.AddWithValue("@injectionAsSeparateBlock", data.injectionAsSeparateBlock);
            cm.Parameters.AddWithValue("@isdaignosisshow", data.isdaignosisshow);
            cm.Parameters.AddWithValue("@foundStatment", data.foundStatment);
            cm.Parameters.AddWithValue("@notfoundStatment", data.notfoundStatment);
            cm.Parameters.AddWithValue("@header_template", data.header_template);
            cm.Parameters.AddWithValue("@font_family", data.font_family);
            cm.Parameters.AddWithValue("@font_size", data.font_size);
            cm.Parameters.AddWithValue("@diagcervialbulge_comma", data.diagcervialbulge_comma);
            cm.Parameters.AddWithValue("@diagthoracicbulge_comma", data.diagthoracicbulge_comma);
            cm.Parameters.AddWithValue("@diaglumberbulge_comma", data.diaglumberbulge_comma);
            cm.Parameters.AddWithValue("@diagleftshoulder_comma", data.diagleftshoulder_comma);
            cm.Parameters.AddWithValue("@diagrightshoulder_comma", data.diagrightshoulder_comma);
            cm.Parameters.AddWithValue("@diagleftknee_comma", data.diagleftknee_comma);
            cm.Parameters.AddWithValue("@diagrightknee_comma", data.diagrightknee_comma);
            cm.Parameters.AddWithValue("@diagbrain_comma", data.diagbrain_comma);
            cm.Parameters.AddWithValue("@other1_comma", data.other1_comma);
            cm.Parameters.AddWithValue("@other2_comma", data.other2_comma);
            cm.Parameters.AddWithValue("@other3_comma", data.other3_comma);
            cm.Parameters.AddWithValue("@other4_comma", data.other4_comma);
            cm.Parameters.AddWithValue("@other5_comma", data.other5_comma);
            cm.Parameters.AddWithValue("@other6_comma", data.other6_comma);
            cm.Parameters.AddWithValue("@other7_comma", data.other7_comma);
            Execute(cm);

        }
        public void Update(tbl_settings data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_settings SET
		page_size=@page_size,
        dateformat=@dateformat,
		location=@location,
        pageBreakForInjection=@pageBreakForInjection,
        injectionAsSeparateBlock=@injectionAsSeparateBlock,
isdaignosisshow=@isdaignosisshow,
      foundStatment=@foundStatment,
      header_template=@header_template,
notfoundStatment=@notfoundStatment,
font_family=@font_family,
font_size=@font_size,
show_preop=@show_preop,
gait_default=@gait_default,
fu_default=@fu_default,
show_postop=@show_postop,
sign_content=@sign_content,
casetype=@casetype,
table_border=@table_border,
diagcervialbulge_comma=@diagcervialbulge_comma,
diagthoracicbulge_comma=@diagthoracicbulge_comma,
diaglumberbulge_comma=@diaglumberbulge_comma,
diagleftshoulder_comma=@diagleftshoulder_comma,
diagrightshoulder_comma=@diagrightshoulder_comma,
diagleftknee_comma=@diagleftknee_comma,
diagrightknee_comma=@diagrightknee_comma,
diagbrain_comma=@diagbrain_comma,
other1_comma=@other1_comma,
other2_comma=@other2_comma,
other3_comma=@other3_comma,
other4_comma=@other4_comma,
other5_comma=@other5_comma,
other6_comma=@other6_comma,
other7_comma=@other7_comma 
			where cmp_id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.cmp_id);
            cm.Parameters.AddWithValue("@page_size", data.page_size);
            cm.Parameters.AddWithValue("@location", data.location);
            cm.Parameters.AddWithValue("@dateformat", data.dateformat);
            cm.Parameters.AddWithValue("@pageBreakForInjection", data.pageBreakForInjection);
            cm.Parameters.AddWithValue("@injectionAsSeparateBlock", data.injectionAsSeparateBlock);
            cm.Parameters.AddWithValue("@isdaignosisshow", data.isdaignosisshow);
            cm.Parameters.AddWithValue("@foundStatment", data.foundStatment);
            cm.Parameters.AddWithValue("@notfoundStatment", data.notfoundStatment);
            cm.Parameters.AddWithValue("@header_template", data.header_template);
            cm.Parameters.AddWithValue("@font_family", data.font_family);
            cm.Parameters.AddWithValue("@font_size", data.font_size);
            cm.Parameters.AddWithValue("@show_preop", data.show_preop);
            cm.Parameters.AddWithValue("@show_postop", data.show_postop);
            cm.Parameters.AddWithValue("@sign_content", data.sign_content);
            cm.Parameters.AddWithValue("@gait_default", data.gait_default);
            cm.Parameters.AddWithValue("@fu_default", data.fu_default);
            cm.Parameters.AddWithValue("@casetype", data.casetype);
            cm.Parameters.AddWithValue("@diagcervialbulge_comma", data.diagcervialbulge_comma);
            cm.Parameters.AddWithValue("@diagthoracicbulge_comma", data.diagthoracicbulge_comma);
            cm.Parameters.AddWithValue("@diaglumberbulge_comma", data.diaglumberbulge_comma);
            cm.Parameters.AddWithValue("@diagleftshoulder_comma", data.diagleftshoulder_comma);
            cm.Parameters.AddWithValue("@diagrightshoulder_comma", data.diagrightshoulder_comma);
            cm.Parameters.AddWithValue("@diagleftknee_comma", data.diagleftknee_comma);
            cm.Parameters.AddWithValue("@diagrightknee_comma", data.diagrightknee_comma);
            cm.Parameters.AddWithValue("@diagbrain_comma", data.diagbrain_comma);
            cm.Parameters.AddWithValue("@other1_comma", data.other1_comma);
            cm.Parameters.AddWithValue("@other2_comma", data.other2_comma);
            cm.Parameters.AddWithValue("@other3_comma", data.other3_comma);
            cm.Parameters.AddWithValue("@other4_comma", data.other4_comma);
            cm.Parameters.AddWithValue("@other5_comma", data.other5_comma);
            cm.Parameters.AddWithValue("@other6_comma", data.other6_comma);
            cm.Parameters.AddWithValue("@other7_comma", data.other7_comma);
            cm.Parameters.AddWithValue("@table_border", data.table_border);

            Execute(cm);
        }
        public void Delete(tbl_settings data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_settings
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.id);
            Execute(cm);
        }
    }
}
