using MySql.Data.MySqlClient;
using System.Linq;
using System.Data;
using MS.Models;
using PainTrax.Services;
using PainTrax.Web.ViewModel;

namespace MS.Services;
public class AppointmentService: ParentService {
	public List<tbl_appointment> GetAll() {
		List<tbl_appointment> dataList = ConvertDataTable<tbl_appointment>(GetData("select * from tbl_appointment")); 
	return dataList;
	}

	public tbl_appointment? GetOne(int appointment_id) {
		DataTable dt = new DataTable();
		 MySqlCommand cm = new MySqlCommand("select * from tbl_appointment where appointment_id=@appointment_id ", conn);
		cm.Parameters.AddWithValue("@appointment_id", appointment_id);
		var datalist = ConvertDataTable<tbl_appointment>(GetData(cm)).FirstOrDefault(); 
		return datalist;
	}

	public void Insert(tbl_appointment data) {
		 MySqlCommand cm = new  MySqlCommand(@"INSERT INTO tbl_appointment
		(patientIE_id,patient_id,location_id,appointmentDate,appointmentStart,appointmentEnd,appointmentNote,status_id)Values
				(@patientIE_id,@patient_id,@location_id,@appointmentDate,@appointmentStart,@appointmentEnd,@appointmentNote,@status_id)",conn);
		cm.Parameters.AddWithValue("@patientIE_id", data.patientIE_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@location_id", data.location_id);
		cm.Parameters.AddWithValue("@appointmentDate", data.appointmentDate);
		cm.Parameters.AddWithValue("@appointmentStart", data.appointmentStart);
		cm.Parameters.AddWithValue("@appointmentEnd", data.appointmentEnd);
		cm.Parameters.AddWithValue("@appointmentNote", data.appointmentNote);
		cm.Parameters.AddWithValue("@status_id", data.status_id);
  	Execute(cm);
}
	public void InsertWithProvider(AppointmentVM data)
	{
		MySqlCommand cm = new MySqlCommand(@"SP_InsertAppointment", conn);
		cm.CommandType = CommandType.StoredProcedure;
		cm.Parameters.AddWithValue("@p_id", data.patient_id);
		cm.Parameters.AddWithValue("@p_lid", data.location_id);
		cm.Parameters.AddWithValue("@p_newDate", data.appointmentDate);
		cm.Parameters.AddWithValue("@p_start", data.appointmentStart);
		cm.Parameters.AddWithValue("@p_end", data.appointmentEnd);
		cm.Parameters.AddWithValue("@p_note", data.appointmentNote);
		cm.Parameters.AddWithValue("@p_statusid", data.status_id);
		cm.Parameters.AddWithValue("@p_providerid", data.providers);
		Execute(cm);
	}
	public void UpdateWithProvider(AppointmentVM data)
	{
		MySqlCommand cm = new MySqlCommand(@"SP_UpdateAppointment", conn);
		cm.CommandType = CommandType.StoredProcedure;
		cm.Parameters.AddWithValue("@p_appid", data.appointment_id);
		cm.Parameters.AddWithValue("@p_lid", data.location_id);
		cm.Parameters.AddWithValue("@p_newDate", data.appointmentDate);
		cm.Parameters.AddWithValue("@p_start", data.appointmentStart);
		cm.Parameters.AddWithValue("@p_end", data.appointmentEnd);
		cm.Parameters.AddWithValue("@p_note", data.appointmentNote);
		cm.Parameters.AddWithValue("@p_providerid", data.providers);
		Execute(cm);
	}


	public void Update(tbl_appointment data) {
		 MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_appointment SET
		patientIE_id=@patientIE_id,
		patient_id=@patient_id,
		location_id=@location_id,
		appointmentDate=@appointmentDate,
		appointmentStart=@appointmentStart,
		appointmentEnd=@appointmentEnd,
		appointmentNote=@appointmentNote,
		status_id=@status_id		where appointment_id=@appointment_id",conn);
		cm.Parameters.AddWithValue("@appointment_id", data.appointment_id);
		cm.Parameters.AddWithValue("@patientIE_id", data.patientIE_id);
		cm.Parameters.AddWithValue("@patient_id", data.patient_id);
		cm.Parameters.AddWithValue("@location_id", data.location_id);
		cm.Parameters.AddWithValue("@appointmentDate", data.appointmentDate);
		cm.Parameters.AddWithValue("@appointmentStart", data.appointmentStart);
		cm.Parameters.AddWithValue("@appointmentEnd", data.appointmentEnd);
		cm.Parameters.AddWithValue("@appointmentNote", data.appointmentNote);
		cm.Parameters.AddWithValue("@status_id", data.status_id);
	Execute(cm);
}
	public void UpdateStatus(tbl_appointment data)
	{
		MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_appointment SET
		status_id=@status_id		where appointment_id=@appointment_id", conn);
		cm.Parameters.AddWithValue("@appointment_id", data.appointment_id);
		cm.Parameters.AddWithValue("@status_id", data.status_id);
		Execute(cm);
	}
	public void Delete(tbl_appointment data) {
		 MySqlCommand cm = new MySqlCommand(@"DELETE FROM tbl_appointment
		where appointment_id=@appointment_id",conn);
		cm.Parameters.AddWithValue("@appointment_id", data.appointment_id);
	Execute(cm);
}
	public void Transfer(TransferVM data)
    {
		MySqlCommand cm = new MySqlCommand(@"UPDATE tbl_appointment SET
		appointmentDate=@ToDate where appointmentDate=@FromDate and location_id=@Location_Id", conn);
		cm.Parameters.AddWithValue("@ToDate", data.ToDate);
		cm.Parameters.AddWithValue("@FromDate", data.FromDate);
		cm.Parameters.AddWithValue("@Location_Id", data.Location_Id);
		Execute(cm);
	}



}