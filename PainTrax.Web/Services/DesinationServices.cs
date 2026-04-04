using MySql.Data.MySqlClient;
using PainTrax.Services;

using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PainTrax.Web.Services
{
	public class DesinationServices : ParentService
	{
		// AppConstant appConstant = new AppConstant(); 

		public DesinationServices()
		{
		}


		public List<tbl_designation> GetAll(string cnd = "")
		{
			//AppConstant.WriteLogFile("Hello");
			try
			{
				string query = "select * from tbl_designation where 1=1 ";

				if (!string.IsNullOrEmpty(query))
					query = query + cnd;

				List<tbl_designation> dataList = ConvertDataTable<tbl_designation>(GetData(query));
				return dataList;
			}
			catch (Exception ex)
			{
				AppConstant.WriteLogFile(ex.Message);
			}
			return null;

		}
		public tbl_designation? GetOne(tbl_designation data)
		{
			DataTable dt = new DataTable();
			MySqlCommand cm = new MySqlCommand("select * from tbl_designation where id=@id ", conn);
			cm.Parameters.AddWithValue("@id", data.id);
			var datalist = ConvertDataTable<tbl_designation>(GetData(cm)).FirstOrDefault();
			return datalist;
		}
		public int Insert(tbl_designation data)
		{
			var isExists = this.IsExists(data);

			if (isExists == false)
			{
				MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_designation
		(title,code,cmp_id,createddate,createdby)Values
				(@title,@code,@cmp_id,@createddate,@createdby);select @@identity", conn);
				cm.Parameters.AddWithValue("@title", data.title);
				cm.Parameters.AddWithValue("@code", data.code);
				cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
				cm.Parameters.AddWithValue("@createddate", data.createddate);
				cm.Parameters.AddWithValue("@createdby", data.createdby);

				return ExecuteScalar(cm);
			}
			else
			{
				return -1;
			}

		}
		public void Update(tbl_designation data)
		{
			MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_designation SET
		title=@title,	
		code=@code,
		updateddate=@updateddate,
		updatedby=@updatedby
		where id=@id", conn);
			cm.Parameters.AddWithValue("@id", data.id);
			cm.Parameters.AddWithValue("@title", data.title);
            cm.Parameters.AddWithValue("@code", data.code);
            cm.Parameters.AddWithValue("@updateddate", data.updateddate);
			cm.Parameters.AddWithValue("@updatedby", data.updatedby);

			Execute(cm);
		}
		public void Delete(tbl_designation data)
		{
			MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_designation
		where id=@id", conn);
			cm.Parameters.AddWithValue("@id", data.id);
			Execute(cm);
		}

		public bool IsExists(tbl_designation model)
		{
			var data = this.GetAll(" and title='" + model.title + "' and cmp_id=" + model.cmp_id + "");
			if (data.Count > 0)
				return true;
			else
				return false;
		}
	}
}
