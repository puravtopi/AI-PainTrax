using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace PainTrax.Web.Helper
{
    public class EnumHelper
    {
        public enum CaseType
        {
            [Display(Name = "NF")]
            NF = 1,
            [Display(Name = "MM")]
            MM = 2,
            [Display(Name = "Lien")]
            Lien = 3,
            [Display(Name = "WC")]
            WC = 4,
            [Display(Name = "OON")]
            OON = 5,
            [Display(Name = "LOP")]
            LOP = 6
        }

        public enum VisitType
        {
            [Display(Name = "FU")]
            FU = 1,
            [Display(Name = "Postop FU")]
            ProFU = 2,
            [Display(Name = "Preop H and P")]
            PriFU = 3,
            [Display(Name = "Tele FU")]
            TeleFU = 4
        }

        public enum Gender
        {
            [Display(Name = "")]
            na = 0,
            [Display(Name = "Male")]
            male = 1,
            [Display(Name = "Female")]
            female = 2,
            [Display(Name = "Other")]
            Other = 3

        }

        public enum MrMrs
        {
            [Display(Name = "Mr.")]
            male = 1,
            [Display(Name = "Ms.")]
            female = 2,
            [Display(Name = "Other")]
            Other = 3

        }

        public enum Study1
        {
            [Display(Name = "")]
            NA = 0,
            [Display(Name = "MRI")]
            MRI = 1,
            [Display(Name = "CT-Scan")]
            CTScan = 2,
            [Display(Name = "X-Rays")]
            XRays = 3,
            [Display(Name = "TM Flow")]
            TMFlow = 4
        }

        public enum Study2
        {
            [Display(Name = "")]
            NA = 0,
            [Display(Name = "MRI")]
            MRI = 1,
            [Display(Name = "CT-Scan")]
            CTScan = 2,
            [Display(Name = "X-Rays")]
            XRays = 3,
            [Display(Name = "UE NCV/EMG")]
            UENCV_EMG = 4,
            [Display(Name = "LE NCV/EMG")]
            LENCV_EMG = 5,
            [Display(Name = "TM Flow")]
            TMFlow = 6
        }

        public enum StudyComma
        {
            [Display(Name = "")]
            NA = 0,
            [Display(Name = ":")]
            MRI = 1,
            [Display(Name = "reveals")]
            CTScan = 2,
           
        }

        public static string GetDisplayName(Enum enumValue)
        {
            return enumValue.GetType()
                            .GetMember(enumValue.ToString())
                            .First()
                            .GetCustomAttribute<DisplayAttribute>()
                            .GetName();
        }
        //public static string GetDisplayName(this Enum enumValue)
        //{
        //    return enumValue
        //        .GetType()
        //        .GetMember(enumValue.ToString())
        //        .First()
        //        .GetCustomAttribute<DisplayAttribute>()?
        //        .GetName() ?? enumValue.ToString();
        //}

        //public static IEnumerable<SelectListItem> GetEnumSelectListWithDisplayNames<T>() where T : Enum
        //{
        //    return Enum.GetValues(typeof(T)).Cast<T>().Select(e => new SelectListItem
        //    {
        //        Text = e.GetDisplayName(),
        //        Value = e.GetDisplayName() // <--- This is key: use display name as value
        //    });
        //}
    }
}
