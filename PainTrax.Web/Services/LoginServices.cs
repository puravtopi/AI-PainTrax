using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.ViewModel;
using System.Data;

namespace PainTrax.Web.Services
{
    public class LoginServices : ParentService
    {

        public ServiceResponse<vm_cm_user> CompanyLogin(LoginVM data)
        {
            // var PassWord = EncryptionHelper.Decrypt(data.PassWord);
            ServiceResponse<vm_cm_user> respones = new ServiceResponse<vm_cm_user>();

            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("select * from vm_cm_user where uname=@uname and client_code=@client_code ", conn);

                cm.Parameters.AddWithValue("@uname", data.UserName);
                cm.Parameters.AddWithValue("@client_code", data.ClientCode);
                var datalist = ConvertDataTable<vm_cm_user>(GetData(cm)).FirstOrDefault();


                if (datalist != null)
                {
                    datalist.password = EncryptionHelper.Decrypt(datalist.password);

                    if (datalist.password == data.PassWord)
                    {
                        respones.Model = datalist;
                        respones.Success = true;

                    }
                    else
                    {
                        respones.Success = false;
                    }

                }
            }
            catch (Exception ex)
            {
                respones.Success = false;
            }
            return respones;
        }

        public ServiceResponse<tbl_admin> AdminLogin(LoginVM data)
        {

            ServiceResponse<tbl_admin> respones = new ServiceResponse<tbl_admin>();

            try
            {
                DataTable dt = new DataTable();
                MySqlCommand cm = new MySqlCommand("select * from tbl_admin where uname=@uname ", conn);
                cm.Parameters.AddWithValue("@uname", data.UserName);
                var datalist = ConvertDataTable<tbl_admin>(GetData(cm)).FirstOrDefault();

                if (datalist != null)
                {

                    if (datalist.password == data.PassWord)
                    {
                        respones.Model = datalist;
                        respones.Success = true;

                    }
                    else
                    {
                        respones.Success = false;
                    }

                }
            }
            catch (Exception ex)
            {
                respones.Success = false;
            }
            return respones;
        }
    }
}
