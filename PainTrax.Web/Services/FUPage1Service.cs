using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;

namespace MS.Services;
public class FUPage1Service : ParentService
{
    public List<tbl_fu_page1> GetAll()
    {
        List<tbl_fu_page1> dataList = ConvertDataTable<tbl_fu_page1>(GetData("select * from tbl_fu_page1"));
        return dataList;
    }

    public tbl_fu_page1? GetOne(int id)
    {
        DataTable dt = new DataTable();
        MySqlCommand cm = new MySqlCommand("select * from tbl_fu_page1 where fu_id=@id ", conn);
        cm.Parameters.AddWithValue("@id", id);
        var datalist = ConvertDataTable<tbl_fu_page1>(GetData(cm)).FirstOrDefault();
        return datalist;
    }

    public void Insert(tbl_fu_page1 data)
    {
        MySqlCommand cm = new MySqlCommand(@"INSERT INTO tbl_fu_page1
		(history,bodypart,cc,pe,rom,note,plan,assessment,dd,work_status,ie_id,daignosis_desc,daignosis_delimit,cmp_id,patient_id,occupation,fu_id,psh,pmh,allergies,medication,family_history,social_history,appt_reason,impairment_rating,vital,poc_assesment)Values
				(@history,@bodypart,@cc,@pe,@rom,@note,@plan,@assessment,@dd,@work_status,@ie_id,@daignosis_desc,@daignosis_delimit,@cmp_id,@patient_id,@occupation,@fu_id,@psh,@pmh,@allergies,@medication,@family_history,@social_history,@appt_reason,@impairment_rating,@vital,@poc_assesment)", conn);

        cm.Parameters.AddWithValue("@bodypart", data.bodypart);
        cm.Parameters.AddWithValue("@cc", data.cc);
        cm.Parameters.AddWithValue("@pe", data.pe);
        cm.Parameters.AddWithValue("@rom", data.rom);
        cm.Parameters.AddWithValue("@note", data.note);
        cm.Parameters.AddWithValue("@plan", data.plan);
        cm.Parameters.AddWithValue("@assessment", data.assessment);
        cm.Parameters.AddWithValue("@dd", data.dd);
        cm.Parameters.AddWithValue("@work_status", data.work_status);
        cm.Parameters.AddWithValue("@ie_id", data.ie_id);
        cm.Parameters.AddWithValue("@daignosis_desc", data.daignosis_desc);
        cm.Parameters.AddWithValue("@daignosis_delimit", data.daignosis_delimit);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@occupation", data.occupation);
        cm.Parameters.AddWithValue("@fu_id", data.fu_id);
        cm.Parameters.AddWithValue("@history", data.history);
        cm.Parameters.AddWithValue("@psh", data.psh);
        cm.Parameters.AddWithValue("@pmh", data.pmh);
        cm.Parameters.AddWithValue("@allergies", data.allergies);
        cm.Parameters.AddWithValue("@medication", data.medication);
        cm.Parameters.AddWithValue("@social_history", data.social_history);
        cm.Parameters.AddWithValue("@family_history", data.family_history);
        cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
        cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);
        cm.Parameters.AddWithValue("@vital", data.vital);
        cm.Parameters.AddWithValue("@poc_assesment", data.poc_assesment);
        Execute(cm);
    }
    public void Update(tbl_fu_page1 data)
    {
        MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_fu_page1 SET
		
		bodypart=@bodypart,
		cc=@cc,
        history=@history,
		pe=@pe,
		rom=@rom,
		note=@note,
		plan=@plan,
		assessment=@assessment,
		dd=@dd,
		work_status=@work_status,
		ie_id=@ie_id,
		daignosis_desc=@daignosis_desc,
		daignosis_delimit=@daignosis_delimit,
		cmp_id=@cmp_id,
		patient_id=@patient_id,
		occupation=@occupation,
        psh=@psh,
        medication=@medication,
        family_history=@family_history,
        social_history=@social_history,
        allergies=@allergies,
        impairment_rating=@impairment_rating,
vital=@vital,
poc_assesment=@poc_assesment,
        appt_reason=@appt_reason,
        pmh=@pmh 
        where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);

        cm.Parameters.AddWithValue("@bodypart", data.bodypart);
        cm.Parameters.AddWithValue("@cc", data.cc);
        cm.Parameters.AddWithValue("@pe", data.pe);
        cm.Parameters.AddWithValue("@rom", data.rom);
        cm.Parameters.AddWithValue("@note", data.note);
        cm.Parameters.AddWithValue("@plan", data.plan);
        cm.Parameters.AddWithValue("@assessment", data.assessment);
        cm.Parameters.AddWithValue("@dd", data.dd);
        cm.Parameters.AddWithValue("@work_status", data.work_status);
        cm.Parameters.AddWithValue("@ie_id", data.ie_id);
        cm.Parameters.AddWithValue("@daignosis_desc", data.daignosis_desc);
        cm.Parameters.AddWithValue("@daignosis_delimit", data.daignosis_delimit);
        cm.Parameters.AddWithValue("@cmp_id", data.cmp_id);
        cm.Parameters.AddWithValue("@patient_id", data.patient_id);
        cm.Parameters.AddWithValue("@occupation", data.occupation);
        cm.Parameters.AddWithValue("@history", data.history);
        cm.Parameters.AddWithValue("@psh", data.psh);
        cm.Parameters.AddWithValue("@pmh", data.pmh);
        cm.Parameters.AddWithValue("@allergies", data.allergies);
        cm.Parameters.AddWithValue("@medication", data.medication);
        cm.Parameters.AddWithValue("@social_history", data.social_history);
        cm.Parameters.AddWithValue("@family_history", data.family_history);
        cm.Parameters.AddWithValue("@appt_reason", data.appt_reason);
        cm.Parameters.AddWithValue("@impairment_rating", data.impairment_rating);
        cm.Parameters.AddWithValue("@vital", data.vital);
        cm.Parameters.AddWithValue("@poc_assesment", data.poc_assesment);
        Execute(cm);

        cm = new MySqlCommand(@"UPDATE tbl_ie_page1 SET
		
		psh=@psh,
		pmh=@pmh,
allergies=@allergies,
medication=@medication,
social_history=@social_history,
family_history=@family_history
        where ie_id=@ie_id", conn);
        cm.Parameters.AddWithValue("@id", data.id);

        
        cm.Parameters.AddWithValue("@psh", data.psh);
        cm.Parameters.AddWithValue("@pmh", data.pmh);
        cm.Parameters.AddWithValue("@ie_id", data.ie_id);
        
        Execute(cm);
    }
    public void Delete(tbl_fu_page1 data)
    {
        MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_fu_page1
		where id=@id", conn);
        cm.Parameters.AddWithValue("@id", data.id);
        Execute(cm);
    }

}