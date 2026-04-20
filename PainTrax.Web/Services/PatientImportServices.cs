using Microsoft.CodeAnalysis;
using MS.Services;
using MySql.Data.MySqlClient;
using PainTrax.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Reflection;

namespace PainTrax.Web.Services
{
    public class PatientImportServices : ParentService
    {
        #region local variables
        PatientIEService patientIEService = new PatientIEService();
        FUPage1Service page1FUService = new FUPage1Service();
        #endregion


        public List<PatientImportReportVM> GetPatientImportReport(string cnd)
        {


            string query = "SELECT " +
    "`id`, " +
    "`last_name`, " +
    "`first_name`, " +
    "`dob` AS dob, " +
    "`sex`, " +
    "`address`, " +
    "`phone`, " +
    "`ssn`, " +
    "`employer_company`, " +
    "`employer_address`, " +
    "`emergency_contact`, " +
    "`work_phone`, " +
    "`date_of_accident` AS doa, " +
    "`condition_related_to`, " +
    "`insurance_company`, " +
    "`insurance_address`, " +
    "`insurance_phone`, " +
    "`claim_number`, " +
    "`claim_address`, " +
    "`nf2`, " +
    "`policy_number`, " +
    "`policy_holder`, " +
    "`wcb_number`, " +
    "`carrier_case_number`, " +
    "`policy_adjuster`, " +
    "`attorney`, " +
    "`firm_name`, " +
    "`attorney_address`, " +
    "`attorney_phone`, " +
    "`attorney_fax`, " +
    "`imported_at` AS imported_at, " +
    "`cmpy_id`, " +
    "`loc_id` " +
    "FROM `tbl_Importpatients` ";

            if (!string.IsNullOrEmpty(cnd))
            {
                query = query + " " + cnd;
            }

            MySqlCommand cm = new MySqlCommand(query, conn);

            //var datalist = ConvertDataTable<PatientImportReportVM>(GetData(cm));
            var datalist = DataTableExtensions.ConvertDataTable<PatientImportReportVM>(GetData(cm));
            return datalist;
        }

        public List<string> GetPatientImportReportDate(string cmpid)
        {
            //string query = "SELECT DISTINCT DATE_FORMAT(pd.Scheduled, '%m/%d/%Y') AS Scheduled FROM tbl_procedures_details pd WHERE pd.Scheduled IS NOT NULL and pd.cmp_id=" + cmpid + " AND pd.Scheduled > CURDATE() ORDER BY pd.Scheduled desc";

            string query = "SELECT " +
    "`id`, " +
    "`last_name`, " +
    "`first_name`, " +
    "DATE_FORMAT(`dob`, '%m/%d/%Y') AS dob, " +
    "`sex`, " +
    "`address`, " +
    "`phone`, " +
    "`ssn`, " +
    "`employer_company`, " +
    "`employer_address`, " +
    "`emergency_contact`, " +
    "`work_phone`, " +
    "DATE_FORMAT(`date_of_accident`, '%m/%d/%Y') AS date_of_accident, " +
    "`condition_related_to`, " +
    "`insurance_company`, " +
    "`insurance_address`, " +
    "`insurance_phone`, " +
    "`claim_number`, " +
    "`claim_address`, " +
    "`nf2`, " +
    "`policy_number`, " +
    "`policy_holder`, " +
    "`wcb_number`, " +
    "`carrier_case_number`, " +
    "`policy_adjuster`, " +
    "`attorney`, " +
    "`firm_name`, " +
    "`attorney_address`, " +
    "`attorney_phone`, " +
    "`attorney_fax`, " +
    "DATE_FORMAT(`imported_at`, '%m/%d/%Y %H:%i:%s') AS imported_at, " +
    "`cmpy_id` " +
    "FROM `tbl_Importpatients` " +
    "WHERE `cmpy_id` = " + cmpid + " " +
    "ORDER BY `last_name` ASC, `first_name` ASC";

            MySqlCommand cm = new MySqlCommand(query, conn);

            var datalist = GetData(cm);


            List<string> list = datalist.AsEnumerable()
                             .Select(row => row["imported_at"].ToString())
                             .ToList();


            //        List<PatientImportReportVM> list = datalist.AsEnumerable()
            //.Select(row => new PatientImportReportVM
            //{
            //    id = row["id"] == DBNull.Value ? 0 : Convert.ToInt32(row["id"]),
            //    last_name = row["last_name"] == DBNull.Value ? null : row["last_name"].ToString(),
            //    first_name = row["first_name"] == DBNull.Value ? null : row["first_name"].ToString(),
            //    dob = row["dob"] == DBNull.Value ? null : Convert.ToDateTime(row["dob"]),
            //    sex = row["sex"] == DBNull.Value ? null : row["sex"].ToString(),
            //    address = row["address"] == DBNull.Value ? null : row["address"].ToString(),
            //    phone = row["phone"] == DBNull.Value ? null : row["phone"].ToString(),
            //    ssn = row["ssn"] == DBNull.Value ? null : row["ssn"].ToString(),
            //    employer_company = row["employer_company"] == DBNull.Value ? null : row["employer_company"].ToString(),
            //    employer_address = row["employer_address"] == DBNull.Value ? null : row["employer_address"].ToString(),
            //    emergency_contact = row["emergency_contact"] == DBNull.Value ? null : row["emergency_contact"].ToString(),
            //    work_phone = row["work_phone"] == DBNull.Value ? null : row["work_phone"].ToString(),
            //    date_of_accident = row["date_of_accident"] == DBNull.Value ? null : Convert.ToDateTime(row["date_of_accident"]),
            //    condition_related_to = row["condition_related_to"] == DBNull.Value ? null : row["condition_related_to"].ToString(),
            //    insurance_company = row["insurance_company"] == DBNull.Value ? null : row["insurance_company"].ToString(),
            //    insurance_address = row["insurance_address"] == DBNull.Value ? null : row["insurance_address"].ToString(),
            //    insurance_phone = row["insurance_phone"] == DBNull.Value ? null : row["insurance_phone"].ToString(),
            //    claim_number = row["claim_number"] == DBNull.Value ? null : row["claim_number"].ToString(),
            //    claim_address = row["claim_address"] == DBNull.Value ? null : row["claim_address"].ToString(),
            //    nf2 = row["nf2"] == DBNull.Value ? null : row["nf2"].ToString(),
            //    policy_number = row["policy_number"] == DBNull.Value ? null : row["policy_number"].ToString(),
            //    policy_holder = row["policy_holder"] == DBNull.Value ? null : row["policy_holder"].ToString(),
            //    wcb_number = row["wcb_number"] == DBNull.Value ? null : row["wcb_number"].ToString(),
            //    carrier_case_number = row["carrier_case_number"] == DBNull.Value ? null : row["carrier_case_number"].ToString(),
            //    policy_adjuster = row["policy_adjuster"] == DBNull.Value ? null : row["policy_adjuster"].ToString(),
            //    attorney = row["attorney"] == DBNull.Value ? null : row["attorney"].ToString(),
            //    firm_name = row["firm_name"] == DBNull.Value ? null : row["firm_name"].ToString(),
            //    attorney_address = row["attorney_address"] == DBNull.Value ? null : row["attorney_address"].ToString(),
            //    attorney_phone = row["attorney_phone"] == DBNull.Value ? null : row["attorney_phone"].ToString(),
            //    attorney_fax = row["attorney_fax"] == DBNull.Value ? null : row["attorney_fax"].ToString(),
            //    imported_at = row["imported_at"] == DBNull.Value ? DateTime.Now : Convert.ToDateTime(row["imported_at"]),
            //    cmpy_id = row["cmpy_id"] == DBNull.Value ? null : Convert.ToInt32(row["cmpy_id"]),
            //})
            //.ToList();


            return list;
        }

    }
}
public static class DataTableExtensions
{
    public static List<T> ConvertDataTable<T>(DataTable dt)
    {
        List<T> data = new List<T>();
        foreach (DataRow row in dt.Rows)
        {
            T item = GetItem<T>(row);
            data.Add(item);
        }
        return data;
    }

    private static T GetItem<T>(DataRow dr)
    {
        Type temp = typeof(T);
        T obj = Activator.CreateInstance<T>();

        foreach (DataColumn column in dr.Table.Columns)
        {
            foreach (PropertyInfo pro in temp.GetProperties())
            {
                if (pro.Name.Equals(column.ColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    var value = dr[column.ColumnName];

                    if (value == DBNull.Value || value == null)
                    {
                        pro.SetValue(obj, null, null);
                        break;
                    }

                    Type propertyType = Nullable.GetUnderlyingType(pro.PropertyType) ?? pro.PropertyType;

                    try
                    {
                        // Case 1: SQL returned DateTime, ViewModel wants string
                        if (propertyType == typeof(string) && value is DateTime dateValue)
                        {
                            pro.SetValue(obj, dateValue.ToString("yyyy-MM-dd"), null);
                        }
                        // ✅ Case 2: SQL returned string (via DATE_FORMAT), ViewModel wants DateTime
                        else if (propertyType == typeof(DateTime) && value is string dateStr)
                        {
                            if (string.IsNullOrWhiteSpace(dateStr))
                            {
                                pro.SetValue(obj, null, null);
                            }
                            else if (DateTime.TryParse(dateStr, out DateTime parsedDate))
                            {
                                // Guard against sentinel value "01/01/0001"
                                pro.SetValue(obj, parsedDate == DateTime.MinValue ? (object)null : parsedDate, null);
                            }
                            else
                            {
                                pro.SetValue(obj, null, null);
                            }
                        }
                        else
                        {
                            object safeValue = Convert.ChangeType(value, propertyType);
                            pro.SetValue(obj, safeValue, null);
                        }
                    }
                    catch
                    {
                        pro.SetValue(obj, null, null);
                    }
                    break;
                }
            }
        }
        return obj;
    }
}
