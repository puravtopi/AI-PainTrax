using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class LogService : ParentService
    {
        /*private readonly ILogger<logService> _logger;
        public logService()
        {
        }
        public logService(ILogger<logService> logger)
        {
            _logger = logger;
        }*/
        public List<tbl_log> GetAll(string cnd = "")
        {
            string query = "select * from tbl_log where 1 = 1";
            if (!string.IsNullOrEmpty(query))
                query = query + cnd;
            List<tbl_log> dataList = ConvertDataTable<tbl_log>(GetData(query));
            return dataList;
        }
        public int Insert(tbl_log data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_log(CreatedDate,CreatedBy,Message)Values
            (@CreatedDate,@CreatedBy,@Message)", conn);
            cm.Parameters.AddWithValue("@CreatedDate", System.DateTime.Now);
            cm.Parameters.AddWithValue("@CreatedBy", data.CreatedBy);
            cm.Parameters.AddWithValue("@Message", data.Message);
            // var result = ExecuteScalar(cm);
            int result = Execute(cm);
            return result;
        }
        public void Delete(tbl_log data)
        {
            MySqlCommand cm = new MySqlCommand(@" DELETE FROM tbl_log where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }
    }
}
