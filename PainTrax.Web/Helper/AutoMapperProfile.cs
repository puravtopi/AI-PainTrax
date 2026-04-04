using AutoMapper;
using MS.Models;

namespace PainTrax.Web.Helper
{
    public class AutoMapperProfile : Profile
    {
        public AutoMapperProfile()
        {
            //source mapping to destination
            //CreateMap<TblCompany, CompanyVM>().ReverseMap();
            //CreateMap<TblDesignation, DesignationVM>().ReverseMap();
            CreateMap<tbl_ie_page1, tbl_fu_page1>().ReverseMap();
            CreateMap<tbl_ie_page2, tbl_fu_page2>().ReverseMap();
            CreateMap<tbl_ie_page3, tbl_fu_page3>().ReverseMap();
            CreateMap<tbl_ie_ne, tbl_fu_ne>().ReverseMap();
            CreateMap<tbl_ie_other, tbl_fu_other>().ReverseMap();
            CreateMap<tbl_ie_comment, tbl_fu_comment>().ReverseMap();
        }
    }
}
