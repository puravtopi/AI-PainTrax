using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class DiagcodesService : ParentService
    {

        private readonly ILogger<DiagcodesService> _logger;

        public DiagcodesService()
        {
        }

        public DiagcodesService(ILogger<DiagcodesService> logger)
        {
            _logger = logger;
        }

        public List<tbl_diagcodes> GetAll(string cnd = "")
        {
            try
            {
                string query = "select * from tbl_diagcodes where 1=1 ";
                query = query + cnd;
                List<tbl_diagcodes> dataList = ConvertDataTable<tbl_diagcodes>(GetData(query));
                return dataList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return null;
            }

        }

        public tbl_diagcodes? GetOne(tbl_diagcodes data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_diagcodes where Id=@id ", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            var datalist = ConvertDataTable<tbl_diagcodes>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public void Insert(tbl_diagcodes data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_diagcodes
		(BodyPart,DiagCode,Description,CreatedDate,CreatedBy,PreSelect,cmp_id,display_order)Values
				(@BodyPart,@DiagCode,@Description,@CreatedDate,@CreatedBy,@PreSelect,@cmp_id,@display_order)", conn);
            cm.Parameters.AddWithValue("@BodyPart", data.BodyPart);
            cm.Parameters.AddWithValue("@DiagCode", data.DiagCode);
            cm.Parameters.AddWithValue("@Description", data.Description);
            cm.Parameters.AddWithValue("@CreatedDate", data.CreatedDate);
            cm.Parameters.AddWithValue("@CreatedBy", data.CreatedBy);
            cm.Parameters.AddWithValue("@PreSelect", data.PreSelect);
            cm.Parameters.AddWithValue("@display_order", data.display_order);

            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void Update(tbl_diagcodes data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_diagcodes SET
		BodyPart=@BodyPart,
		DiagCode=@DiagCode,
		Description=@Description,
		display_order=@display_order,
	    PreSelect=@PreSelect
	    where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            cm.Parameters.AddWithValue("@BodyPart", data.BodyPart);
            cm.Parameters.AddWithValue("@DiagCode", data.DiagCode);
            cm.Parameters.AddWithValue("@Description", data.Description);
            cm.Parameters.AddWithValue("@PreSelect", data.PreSelect);
            cm.Parameters.AddWithValue("@display_order", data.display_order);

            Execute(cm);
        }
        public void Delete(tbl_diagcodes data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_diagcodes
		where Id=@Id", conn);
            cm.Parameters.AddWithValue("@Id", data.Id);
            Execute(cm);
        }

    }
}