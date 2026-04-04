using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class POCConfigService : ParentService
    {
        public List<tbl_pocconfig> GetAll(string cnd = "")
        {
            string query = "select * from tbl_pocconfig where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            var datalist = ConvertDataTable<tbl_pocconfig>(GetData(query)).FirstOrDefault();

            List<tbl_pocconfig> configList = new List<tbl_pocconfig>();
            
            if (datalist == null || string.IsNullOrEmpty(datalist.columns))
            {
                return configList; // return empty list safely
            }

            string[] entries = datalist.columns.Split(',');

            foreach (string entry in entries)
            {
                //  string[] parts = entry.Split(',');

                if (entry.Length >= 1)
                {
                    tbl_pocconfig config = new tbl_pocconfig
                    {
                        id = entry.TrimEnd(),
                        columns = entry.TrimEnd(),
                        //Value = parts[1]
                    };

                    configList.Add(config);
                }
            }


            return configList;
        }
        public List<tbl_pocconfig> GetAllExport(string cnd = "")
        {
            string query = "select * from tbl_pocconfig where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            var datalist = ConvertDataTable<tbl_pocconfig>(GetData(query)).FirstOrDefault();

            List<tbl_pocconfig> configList = new List<tbl_pocconfig>();

            if (datalist == null || string.IsNullOrEmpty(datalist.export_Columns))
            {
                return configList; // return empty list safely
            }

            string[] entries = datalist.export_Columns.Split(',');

            foreach (string entry in entries)
            {
                //  string[] parts = entry.Split(',');

                if (entry.Length >= 1)
                {
                    tbl_pocconfig config = new tbl_pocconfig
                    {
                        id = entry.TrimEnd(),
                        columns = entry.TrimEnd(),
                        export_Columns = entry.TrimEnd(),
                        //Value = parts[1]
                    };

                    configList.Add(config);
                }
            }


            return configList;
        }


        public List<tbl_pocconfig> GetAllone(string cnd = "")
        {
            string query = "select * from tbl_pocconfig where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            var datalist = ConvertDataTable<tbl_pocconfig>(GetData(query)).FirstOrDefault();            

            List<tbl_pocconfig> configList = new List<tbl_pocconfig>();

            
            if (datalist == null || string.IsNullOrEmpty(datalist.columns))
            {
                return configList; // return empty list safely
            }

            string[] entries = datalist.columns.Split(',');

            foreach (string entry in entries)
            {
                //  string[] parts = entry.Split(',');

                if (entry.Length >= 1)
                {
                    tbl_pocconfig config = new tbl_pocconfig
                    {
                        id = entry.TrimEnd(),
                        columns = entry.TrimEnd(),
                        //Value = parts[1]
                    };

                    configList.Add(config);
                }
            }


            return configList;
        }
        public List<tbl_pocconfig> GetAlloneExport(string cnd = "")
        {
            string query = "select * from tbl_pocconfig where 1=1 ";

            if (!string.IsNullOrEmpty(cnd))
                query = query + cnd;

            var datalist = ConvertDataTable<tbl_pocconfig>(GetData(query)).FirstOrDefault();

            List<tbl_pocconfig> configList = new List<tbl_pocconfig>();


            if (datalist == null || string.IsNullOrEmpty(datalist.export_Columns))
            {
                return configList; // return empty list safely
            }

            string[] entries = datalist.export_Columns.Split(',');

            foreach (string entry in entries)
            {
                //  string[] parts = entry.Split(',');

                if (entry.Length >= 1)
                {
                    tbl_pocconfig config = new tbl_pocconfig
                    {
                        id = entry.TrimEnd(),
                        columns = entry.TrimEnd(),
                        export_Columns = entry.TrimEnd(),
                        //Value = parts[1]
                    };

                    configList.Add(config);
                }
            }


            return configList;
        }

        public tbl_pocconfig? GetOne(tbl_pocconfig data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_pocconfig ", conn);
            //cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_pocconfig>(GetData(cm)).FirstOrDefault();
            if (datalist == null)
            {
                return null; // or you could return a new tbl_pocconfig() if preferred
            }
            return datalist;
        }

        public void Insert(tbl_pocconfig data)
        {
            MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_pocconfig
		(id,columns,cmp_id)Values
				(@id,@columns,@cmp_id);select @@identity;", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@columns", data.columns);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);

            //cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            //cm.Parameters.AddWithValue("@createddate", data.createddate);
            //cm.Parameters.AddWithValue("@createdby", data.createdby);
            //cm.Parameters.AddWithValue("@header_template", data.header_template);

            int id = ExecuteScalar(cm);

            //addlocationintogroup(data.cmp_id.Value, data.location, id);
        }
        public void Update(tbl_pocconfig data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_pocconfig SET
            id = @id,		
            columns=@columns where cmp_id=@cmp_id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@columns", data.columns);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }
        public void ExportUpdate(tbl_pocconfig data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_pocconfig SET
            id = @id,		
            export_Columns=@export_Columns  where cmp_id = @cmp_id", conn);

            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@export_Columns", data.export_Columns);
            cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            Execute(cm);
        }


    }
}