using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services

{
    public class ProcedureService : ParentService
    {
        public List<tbl_procedures> GetAll(string cnd = "")
        {
            string query = "select * from tbl_procedures where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;

            List<tbl_procedures> dataList = ConvertDataTable<tbl_procedures>(GetData(query));
            return dataList;
        }

        public tbl_procedures? GetOne(tbl_procedures data)
        {
            DataTable dt = new DataTable();
            MySqlCommand cm = new MySqlCommand("select * from tbl_procedures where id=@id ", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            var datalist = ConvertDataTable<tbl_procedures>(GetData(cm)).FirstOrDefault();
            return datalist;
        }

        public void Insert(tbl_procedures data)
        {
            try
            {
                MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_procedures
		(mcode,bodypart,heading,ccdesc,pedesc,adesc,pdesc,cf,pn,preselect,display_order,position,inhouseproc,inhouseprocbit,haslevel,hasmuscle,hasmedication,bid,tbl_procedures.inout,sides,tbl_procedures.status,hassubprocedure,s_ccdesc,s_pedesc,s_adesc,s_pdesc,e_ccdesc,e_pedesc,e_adesc,e_pdesc,s_heading,e_heading,levelsdefault,sidesdefault,cmp_id,mcode_desc,upload_template,injection_description)Values
				(@mcode,@bodypart,@heading,@ccdesc,@pedesc,@adesc,@pdesc,@cf,@pn,@preselect,@display_order,@position,@inhouseproc,@inhouseprocbit,@haslevel,@hasmuscle,@hasmedication,@bid,@inout,@sides,@status,@hassubprocedure,@s_ccdesc,@s_pedesc,@s_adesc,@s_pdesc,@e_ccdesc,@e_pedesc,@e_adesc,@e_pdesc,@s_heading,@e_heading,@levelsdefault,@sidesdefault,@cmp_id,@mcode_desc,@upload_template,@injection_description)", conn);
                cm.Parameters.AddWithValue("@mcode", data.mcode);
                cm.Parameters.AddWithValue("@bodypart", data.bodypart);
                cm.Parameters.AddWithValue("@heading", data.heading);
                cm.Parameters.AddWithValue("@ccdesc", data.ccdesc);
                cm.Parameters.AddWithValue("@pedesc", data.pedesc);
                cm.Parameters.AddWithValue("@adesc", data.adesc);
                cm.Parameters.AddWithValue("@pdesc", data.pdesc);
                cm.Parameters.AddWithValue("@cf", data.cf);
                cm.Parameters.AddWithValue("@pn", data.pn);
                cm.Parameters.AddWithValue("@preselect", data.preselect);
                cm.Parameters.AddWithValue("@display_order", data.display_order);
                cm.Parameters.AddWithValue("@position", data.position);
                cm.Parameters.AddWithValue("@inhouseproc", data.inhouseproc);
                cm.Parameters.AddWithValue("@inhouseprocbit", data.inhouseprocbit);
                cm.Parameters.AddWithValue("@haslevel", data.haslevel);
                cm.Parameters.AddWithValue("@hasmuscle", data.hasmuscle);
                cm.Parameters.AddWithValue("@hasmedication", data.hasmedication);
                cm.Parameters.AddWithValue("@bid", data.bid);
                cm.Parameters.AddWithValue("@inout", data.inout);
                cm.Parameters.AddWithValue("@sides", data.sides);
                cm.Parameters.AddWithValue("@status", data.status);
                cm.Parameters.AddWithValue("@hassubprocedure", data.hassubprocedure);
                cm.Parameters.AddWithValue("@s_ccdesc", data.s_ccdesc);
                cm.Parameters.AddWithValue("@s_pedesc", data.s_pedesc);
                cm.Parameters.AddWithValue("@s_adesc", data.s_adesc);
                cm.Parameters.AddWithValue("@s_pdesc", data.s_pdesc);
                cm.Parameters.AddWithValue("@e_ccdesc", data.e_ccdesc);
                cm.Parameters.AddWithValue("@e_pedesc", data.e_pedesc);
                cm.Parameters.AddWithValue("@e_adesc", data.e_adesc);
                cm.Parameters.AddWithValue("@e_pdesc", data.e_pdesc);
                cm.Parameters.AddWithValue("@s_heading", data.s_heading);
                cm.Parameters.AddWithValue("@e_heading", data.e_heading);
                cm.Parameters.AddWithValue("@levelsdefault", data.levelsdefault);
                cm.Parameters.AddWithValue("@sidesdefault", data.sidesdefault);
                cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
                cm.Parameters.AddWithValue("@mcode_desc", data.mcode_desc);
                cm.Parameters.AddWithValue("@upload_template", data.upload_template);
                cm.Parameters.AddWithValue("@injection_description", data.injection_description);
                Execute(cm);
            }
            catch (Exception ex)
            {

                string mmm = ex.Message;
            }
        }
        public void Update(tbl_procedures data)
        {
            MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_procedures SET
		mcode=@mcode,
		bodypart=@bodypart,
		heading=@heading,
		ccdesc=@ccdesc,
		pedesc=@pedesc,
		adesc=@adesc,
		pdesc=@pdesc,
		cf=@cf,
		pn=@pn,
		preselect=@preselect,
		display_order=@display_order,
		position=@position,
		inhouseproc=@inhouseproc,
		inhouseprocbit=@inhouseprocbit,
		haslevel=@haslevel,
		hasmuscle=@hasmuscle,
		hasmedication=@hasmedication,
		bid=@bid,
		tbl_procedures.inout=@inout,
		sides=@sides,
		tbl_procedures.status=@status,
		hassubprocedure=@hassubprocedure,
		s_ccdesc=@s_ccdesc,
		s_pedesc=@s_pedesc,
		s_adesc=@s_adesc,
		s_pdesc=@s_pdesc,
		e_ccdesc=@e_ccdesc,
		e_pedesc=@e_pedesc,
		e_adesc=@e_adesc,
		e_pdesc=@e_pdesc,
		s_heading=@s_heading,
		e_heading=@e_heading,
        mcode_desc=@mcode_desc,
		levelsdefault=@levelsdefault,
		sidesdefault=@sidesdefault,
        upload_template=@upload_template,
        injection_description=@injection_description

				where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            cm.Parameters.AddWithValue("@mcode", data.mcode);
            cm.Parameters.AddWithValue("@bodypart", data.bodypart);
            cm.Parameters.AddWithValue("@heading", data.heading);
            cm.Parameters.AddWithValue("@ccdesc", data.ccdesc);
            cm.Parameters.AddWithValue("@pedesc", data.pedesc);
            cm.Parameters.AddWithValue("@adesc", data.adesc);
            cm.Parameters.AddWithValue("@pdesc", data.pdesc);
            cm.Parameters.AddWithValue("@cf", data.cf);
            cm.Parameters.AddWithValue("@pn", data.pn);
            cm.Parameters.AddWithValue("@preselect", data.preselect);
            cm.Parameters.AddWithValue("@display_order", data.display_order);
            cm.Parameters.AddWithValue("@position", data.position);
            cm.Parameters.AddWithValue("@inhouseproc", data.inhouseproc);
            cm.Parameters.AddWithValue("@inhouseprocbit", data.inhouseprocbit);
            cm.Parameters.AddWithValue("@haslevel", data.haslevel);
            cm.Parameters.AddWithValue("@hasmuscle", data.hasmuscle);
            cm.Parameters.AddWithValue("@hasmedication", data.hasmedication);
            cm.Parameters.AddWithValue("@bid", data.bid);
            cm.Parameters.AddWithValue("@inout", data.inout);
            cm.Parameters.AddWithValue("@sides", data.sides);
            cm.Parameters.AddWithValue("@status", data.status);
            cm.Parameters.AddWithValue("@hassubprocedure", data.hassubprocedure);
            cm.Parameters.AddWithValue("@s_ccdesc", data.s_ccdesc);
            cm.Parameters.AddWithValue("@s_pedesc", data.s_pedesc);
            cm.Parameters.AddWithValue("@s_adesc", data.s_adesc);
            cm.Parameters.AddWithValue("@s_pdesc", data.s_pdesc);
            cm.Parameters.AddWithValue("@e_ccdesc", data.e_ccdesc);
            cm.Parameters.AddWithValue("@e_pedesc", data.e_pedesc);
            cm.Parameters.AddWithValue("@e_adesc", data.e_adesc);
            cm.Parameters.AddWithValue("@e_pdesc", data.e_pdesc);
            cm.Parameters.AddWithValue("@s_heading", data.s_heading);
            cm.Parameters.AddWithValue("@e_heading", data.e_heading);
            cm.Parameters.AddWithValue("@levelsdefault", data.levelsdefault);
            cm.Parameters.AddWithValue("@sidesdefault", data.sidesdefault);
            cm.Parameters.AddWithValue("@cmpid", data.cmp_id);
            cm.Parameters.AddWithValue("@mcode_desc", data.mcode_desc);
            cm.Parameters.AddWithValue("@upload_template", data.upload_template);
            cm.Parameters.AddWithValue("@injection_description", data.injection_description);
            Execute(cm);
        }
        public void Delete(tbl_procedures data)
        {
            MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_procedures
		where id=@id", conn);
            cm.Parameters.AddWithValue("@id", data.id);
            Execute(cm);
        }

        public int? GetId(string cnd = "")
        {
            string query = "select id from tbl_procedures where 1=1 ";

            if (!string.IsNullOrEmpty(query))
                query = query + cnd;

            List<tbl_procedures> dataList = ConvertDataTable<tbl_procedures>(GetData(query));
            if (dataList != null)
            {
                if (dataList.Count > 0)
                {
                    return Convert.ToInt32(dataList[0].id);
                }
                else
                    return null;
            }
            else
                return null;
        }

    }
}

