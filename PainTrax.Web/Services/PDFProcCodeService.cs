using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Models;

namespace PainTrax.Web.Services
{
    public class PdfProcCodeService : ParentService
    {
        // ----------------------------------------------------
        // GET ALL
        // ----------------------------------------------------
        public List<tbl_pdfproccode> GetAll(string cnd = "")
        {
            string sql = "SELECT * FROM tbl_pdfproccode where 1=1 ";
            if (!string.IsNullOrEmpty(cnd))
                sql = sql + cnd;

            return ConvertDataTable<tbl_pdfproccode>(GetData(sql));
        }

        // ----------------------------------------------------
        // GET ONE
        // ----------------------------------------------------
        public tbl_pdfproccode? GetOne(int id)
        {
            MySqlCommand cmd = new MySqlCommand(
                "SELECT * FROM tbl_pdfproccode WHERE mcode=@id", conn);

            cmd.Parameters.AddWithValue("@id", id);

            return ConvertDataTable<tbl_pdfproccode>(GetData(cmd))
                   .FirstOrDefault();
        }

        public tbl_pdfproccode? GetMcode(string mcode)
        {
            MySqlCommand cmd = new MySqlCommand(
                "SELECT * FROM tbl_pdfproccode WHERE mcode=@mcode", conn);

            cmd.Parameters.AddWithValue("@mcode", mcode);

            return ConvertDataTable<tbl_pdfproccode>(GetData(cmd))
                   .FirstOrDefault();
        }

        // ----------------------------------------------------
        // INSERT
        // ----------------------------------------------------
        public void Insert(tbl_pdfproccode data)
        {
            MySqlCommand cmd = new MySqlCommand(@"
            INSERT INTO tbl_pdfproccode
            (mcode, mprocedure, cptcodes, icdcodes, specialequ, diagnosis, speequchk, 
             mprocshort, cptcode1, cptcode2, cptcode3, cptcode4,
             icdcode1, icdcode2, icdcode3, icdcode4, cmp_id, cmp_code)
            VALUES
            (@mcode, @mprocedure, @cptcodes, @icdcodes, @specialequ, @diagnosis, @speequchk,
             @mprocshort, @cptcode1, @cptcode2, @cptcode3, @cptcode4,
             @icdcode1, @icdcode2, @icdcode3, @icdcode4, @cmp_id, @cmp_code)
        ", conn);

            cmd.Parameters.AddWithValue("@mcode", data.mcode);
            cmd.Parameters.AddWithValue("@mprocedure", data.mprocedure);
            cmd.Parameters.AddWithValue("@cptcodes", data.cptcodes);
            cmd.Parameters.AddWithValue("@icdcodes", data.icdcodes);
            cmd.Parameters.AddWithValue("@specialequ", data.specialequ);
            cmd.Parameters.AddWithValue("@diagnosis", data.diagnosis);
            cmd.Parameters.AddWithValue("@speequchk", data.speequchk);
            cmd.Parameters.AddWithValue("@mprocshort", data.mprocshort);
            cmd.Parameters.AddWithValue("@cptcode1", data.cptcode1);
            cmd.Parameters.AddWithValue("@cptcode2", data.cptcode2);
            cmd.Parameters.AddWithValue("@cptcode3", data.cptcode3);
            cmd.Parameters.AddWithValue("@cptcode4", data.cptcode4);
            cmd.Parameters.AddWithValue("@icdcode1", data.icdcode1);
            cmd.Parameters.AddWithValue("@icdcode2", data.icdcode2);
            cmd.Parameters.AddWithValue("@icdcode3", data.icdcode3);
            cmd.Parameters.AddWithValue("@icdcode4", data.icdcode4);
            cmd.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cmd.Parameters.AddWithValue("@cmp_code", data.cmp_code);

            Execute(cmd);
        }

        // ----------------------------------------------------
        // UPDATE
        // ----------------------------------------------------
        public void Update(tbl_pdfproccode data)
        {
            MySqlCommand cmd = new MySqlCommand(@"
            UPDATE tbl_pdfproccode SET
                mcode=@mcode,
                mprocedure=@mprocedure,
                cptcodes=@cptcodes,
                icdcodes=@icdcodes,
                specialequ=@specialequ,
                diagnosis=@diagnosis,
                speequchk=@speequchk,
                mprocshort=@mprocshort,
                cptcode1=@cptcode1,
                cptcode2=@cptcode2,
                cptcode3=@cptcode3,
                cptcode4=@cptcode4,
                icdcode1=@icdcode1,
                icdcode2=@icdcode2,
                icdcode3=@icdcode3,
                icdcode4=@icdcode4,
                cmp_id=@cmp_id,
                cmp_code=@cmp_code
            WHERE id=@id
        ", conn);

            cmd.Parameters.AddWithValue("@id", data.id);
            cmd.Parameters.AddWithValue("@mcode", data.mcode);
            cmd.Parameters.AddWithValue("@mprocedure", data.mprocedure);
            cmd.Parameters.AddWithValue("@cptcodes", data.cptcodes);
            cmd.Parameters.AddWithValue("@icdcodes", data.icdcodes);
            cmd.Parameters.AddWithValue("@specialequ", data.specialequ);
            cmd.Parameters.AddWithValue("@diagnosis", data.diagnosis);
            cmd.Parameters.AddWithValue("@speequchk", data.speequchk);
            cmd.Parameters.AddWithValue("@mprocshort", data.mprocshort);
            cmd.Parameters.AddWithValue("@cptcode1", data.cptcode1);
            cmd.Parameters.AddWithValue("@cptcode2", data.cptcode2);
            cmd.Parameters.AddWithValue("@cptcode3", data.cptcode3);
            cmd.Parameters.AddWithValue("@cptcode4", data.cptcode4);
            cmd.Parameters.AddWithValue("@icdcode1", data.icdcode1);
            cmd.Parameters.AddWithValue("@icdcode2", data.icdcode2);
            cmd.Parameters.AddWithValue("@icdcode3", data.icdcode3);
            cmd.Parameters.AddWithValue("@icdcode4", data.icdcode4);
            cmd.Parameters.AddWithValue("@cmp_id", data.cmp_id);
            cmd.Parameters.AddWithValue("@cmp_code", data.cmp_code);

            Execute(cmd);
        }

        // ----------------------------------------------------
        // DELETE
        // ----------------------------------------------------
        public void Delete(int id)
        {
            MySqlCommand cmd = new MySqlCommand(
                "DELETE FROM tbl_pdfproccode WHERE id=@id", conn);

            cmd.Parameters.AddWithValue("@id", id);

            Execute(cmd);
        }
    }

}
