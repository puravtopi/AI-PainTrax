using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;
using System.Data;

namespace PainTrax.Web.Services
{
    public class DictationServices : ParentService
    {


        public tbl_dictation? GetOne(tbl_designation data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_dictation where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_dictation>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public tbl_dictation? GetDictationAudio(int ieId,string type,int fuId=0)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_dictation where ie_id=@ie_id and type=@type", conn);
            cm.Parameters.AddWithValue("@ie_id", ieId);
            cm.Parameters.AddWithValue("@type", type);
            var datalist = ConvertDataTable<tbl_dictation>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public int Insert(tbl_dictation data)
        {


            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_dictation
		where ie_id=@ie_id and type=@type", conn);
            cm.Parameters.AddWithValue("@ie_id", data.ie_id);
            cm.Parameters.AddWithValue("@type", data.type);
            Execute(cm);


             cm = new MySqlCommand(@"INSERT INTO tbl_dictation
		(voice_file,txt_file,type,ie_id,fu_id,cmp_id)Values
				(@voice_file,@txt_file,@type,@ie_id,@fu_id,@cmp_id);select @@identity", conn);
            cm.Parameters.AddWithValue("@voice_file", data.voice_file);
            cm.Parameters.AddWithValue("@txt_file", data.txt_file);
            cm.Parameters.AddWithValue("@type", data.type);
            cm.Parameters.AddWithValue("@ie_id", data.ie_id);
            cm.Parameters.AddWithValue("@fu_id", data.fu_id);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);


            return ExecuteScalar(cm);


        }

        public int Delete(int ie_id,string type)
        {


            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_dictation
		        where ie_id=@ie_id and type=@type", conn);
            cm.Parameters.AddWithValue("@ie_id", ie_id);
            cm.Parameters.AddWithValue("@type", type);
         

            return ExecuteScalar(cm);


        }
    }
}
