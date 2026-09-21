using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class TranscribeService : ParentService
    {

        public int Insert(tbl_transcribe data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_transcribe
		(content,history,cc,plan,subjective,assessment,ie_id,fu_id,length,intake_id,cmp_id,date,file_path)Values
				(@content,@history,@cc,@plan,@subjective,@assessment,@ie_id,@fu_id,@length,@intake_id,@cmp_id,@date,@file_path);select @@identity;", conn);
            cm.Parameters.AddWithValue("@content", data.content);
            cm.Parameters.AddWithValue("@history", data.history);
            cm.Parameters.AddWithValue("@cc", data.cc);
            cm.Parameters.AddWithValue("@plan", data.plan);
            cm.Parameters.AddWithValue("@subjective", data.subjective);
            cm.Parameters.AddWithValue("@assessment", data.assessment);
            cm.Parameters.AddWithValue("@ie_id", data.ie_id);
            cm.Parameters.AddWithValue("@fu_id", data.fu_id);
            cm.Parameters.AddWithValue("@length", data.length);
            cm.Parameters.AddWithValue("@intake_id", data.intake_id);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cm.Parameters.AddWithValue("@date", data.date);
            cm.Parameters.AddWithValue("@file_path", data.file_path);

            var result = ExecuteScalar(cm);
            return result;
        }

        public tbl_transcribe? GetOne(int id)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_transcribe where intake_id=@id ", conn);
            cm.Parameters.AddWithValue("@id", id);
            var datalist = ConvertDataTable<tbl_transcribe>(GetData(cm)).FirstOrDefault();
            return datalist;
        }
    }
}
