using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using PainTrax.Web.ViewModel;
using PainTrax.Web.Models;

namespace MS.Services;
public class SignInSheetService : ParentService
{
    public List<tbl_appointment> GetAll()
    {
        List<tbl_appointment> dataList = ConvertDataTable<tbl_appointment>(GetData("select * from tbl_appointment"));
        return dataList;
    }

    //public List<PatientsByDOE> GetPatientsByDOE(string doe, string Location_ID, string cmpID)
    //{
    //    try
    //    {
    //        if (string.IsNullOrEmpty(Location_ID))
    //        {
    //            Location_ID = "0";

    //        }
    //        MySqlCommand cm = new MySqlCommand("CALL GetPatientsByDOE(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);

    //        //var datalist = (GetData(cm));

    //        // MySqlCommand cm = new MySqlCommand(query, conn);

    //        var datalist = ConvertDataTable<PatientsByDOE>(GetData(cm));
    //        return datalist;





    //    }
    //    catch (Exception ex)
    //    {
    //        return null;
    //    }
    //}

    public DataTable GetPatientsByDOE(string doe, string Location_ID, string cmpID)
    {
        try
        {
            if (string.IsNullOrEmpty(Location_ID))
            {
                Location_ID = "0";
               // Location_ID = null;

            }
            MySqlCommand cm = new MySqlCommand("CALL GetPatientsByDOE(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);

            var datalist = (GetData(cm));

            // MySqlCommand cm = new MySqlCommand(query, conn);

            // var datalist = ConvertDataTable<PatientsByDOE>(GetData(cm));
            return datalist;





        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public List<SI_Report> GetPatientsSIDNL(string doe, string Location_ID, string cmpID)
    {
        try
        {
            //string doe="2025-01-17"; string Location_ID="7"; string cmpID="7";
            if (string.IsNullOrEmpty(Location_ID))
            {
                Location_ID = "0";

            }
            // MySqlCommand cm = new MySqlCommand("CALL GetPatientsSIDNL(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);
            MySqlCommand cm = new MySqlCommand("CALL GetPatientsSIDNLNew(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);
            

            //var datalist = (GetData(cm));

            // MySqlCommand cm = new MySqlCommand(query, conn);

            //var datalist = ConvertDataTable<SI_Report>(GetData(cm));
            List<SI_Report> dataList = ConvertDataTable<SI_Report>(GetData(cm));
            return dataList;

        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public DataTable GetPatientsSIDNLDatatable(string doe, string Location_ID, string cmpID)
    {
        try
        {
            DataTable datalist = new DataTable();
            //string doe = "2025-01-17"; string Location_ID = "7"; string cmpID = "7";
            if (string.IsNullOrEmpty(Location_ID))
            {
                Location_ID = "0";

            }
            //MySqlCommand cm = new MySqlCommand("CALL GetPatientsSIDNL(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);
            MySqlCommand cm = new MySqlCommand("CALL GetPatientsSIDNLNew(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);

            datalist = (GetData(cm));

            // MySqlCommand cm = new MySqlCommand(query, conn);

            //var dataList = ConvertDataTable<SI_Report>(GetData(cm));
            //List<SI_Report> dataList = ConvertDataTable<SI_Report>(GetData(cm));
            return datalist;

        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public List<tbl_attorneys> GetAll(string cnd = "")
    {
        string query = "select * from tbl_attorneys where 1=1 ";

        if (!string.IsNullOrEmpty(query))
            query = query + cnd;

        List<tbl_attorneys> dataList = ConvertDataTable<tbl_attorneys>(GetData(query));
        return dataList;
    }

    public DataTable GetPatientsByDOEDetails(string IEID, string FUID,string type, string cmpID)
    {
        try
        {


            MySqlCommand cm = null;

            // MySqlCommand cm = new MySqlCommand("CALL GetPatientsByDOE(" + cmpID + ",'" + doe + "'," + Location_ID + ")", conn);

            if (type == "IE")
            {
                cm = new MySqlCommand("SELECT ie.doa,ie.doe,pt.fname,pt.lname,pt.dob,tp.provider,lc.location,ie.compensation,p3.diagcervialbulge_date,p3.diagthoracicbulge_date,p3.diaglumberbulge_date,p3.other1_study,p3.other1_date,p3.other2_study,p3.other2_date,p3.other3_study,p3.other3_date,p3.other4_study,p3.other4_date,p3.other5_study,p3.other5_date,p3.other6_study,p3.other6_date,p3.other7_study,p3.other7_date\r\n FROM tbl_patient_ie ie INNER JOIN tbl_locations lc ON lc.id = ie.location_id INNER JOIN tbl_patient pt ON pt.id=ie.patient_id LEFT JOIN tbl_provider tp ON tp.id=ie.provider_id LEFT JOIN tbl_ie_page3 p3 ON p3.ie_id=ie.id  WHERE ie.id=@IEID and pt.cmp_id=@cmpID ", conn);
                cm.Parameters.AddWithValue("@IEID", IEID);
                cm.Parameters.AddWithValue("@cmpID", cmpID);
            }

            if (type == "FU")
            {
                cm = new MySqlCommand("SELECT ie.doa,ie.doe,pt.fname,pt.lname,pt.dob,tp.provider,lc.location,ie.compensation,p3.diagcervialbulge_date,p3.diagthoracicbulge_date,p3.diaglumberbulge_date,p3.other1_study,p3.other1_date,p3.other2_study,p3.other2_date,p3.other3_study,p3.other3_date,p3.other4_study,p3.other4_date,p3.other5_study,p3.other5_date,p3.other6_study,p3.other6_date,p3.other7_study,p3.other7_date \r\nFROM tbl_patient_ie ie INNER JOIN tbl_locations lc ON lc.id = ie.location_id INNER JOIN tbl_patient_fu fu ON fu.patientIE_ID=ie.id INNER JOIN tbl_patient pt ON pt.id=ie.patient_id LEFT JOIN tbl_provider tp ON tp.id=ie.provider_id LEFT JOIN tbl_fu_page3 p3 ON p3.fu_id=fu.id AND p3.ie_id = ie.id  WHERE ie.id=@IEID AND fu.id =@FUID AND pt.cmp_id=@cmpID ", conn);
                cm.Parameters.AddWithValue("@IEID", IEID);
                cm.Parameters.AddWithValue("@FUID", FUID);
                cm.Parameters.AddWithValue("@cmpID", cmpID);
            }

            

            var datalist = (GetData(cm));
            
            return datalist;


        }
        catch (Exception ex)
        {
            return null;
        }
    }



}