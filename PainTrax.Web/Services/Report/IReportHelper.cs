using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Services.Report
{
    public interface IReportHelper
    {
        // General / Summary
        string GetDisabilityStatement(AIIntakeFormModel data);
        string GetHistory(AIIntakeFormModel model);
        string GenerateCervicalSpineReport(AIIntakeFormModel model);
        string GetDiagnosis(string html);
        string GetPalpation(dynamic model);

        string GenerateSocialHistory(AIIntakeFormModel modal);

        // Chief Complaints (CC)
        string NeckCC(AIIntakeFormModel modal);
        string MidbackCC(AIIntakeFormModel modal);
        string LowBackCC(AIIntakeFormModel modal);
        string LeftShoulderCC(AIIntakeFormModel modal);
        string RightShoulderCC(AIIntakeFormModel modal);
        string RightKneeCC(AIIntakeFormModel modal);
        string LeftKneeCC(AIIntakeFormModel modal);
        string RightAnkleCC(AIIntakeFormModel modal);
        string LeftAnkleCC(AIIntakeFormModel modal);
        string RightElbowCC(AIIntakeFormModel modal);
        string LeftElbowCC(AIIntakeFormModel modal);
        string RightHipCC(AIIntakeFormModel modal);
        string LeftHipCC(AIIntakeFormModel modal);
        string RightWristCC(AIIntakeFormModel modal);
        string LeftWristCC(AIIntakeFormModel modal);
        string GenerateOther(string PainScale, List<string> PainTypes, string title);

        // Physical Exam (PE) & AI Content
        string GetCC(AIIntakeFormModel model);
        string GetPE(AIIntakeFormModel model, string path);
        string GetDTR(AIIntakeFormModel model, string path);
        string GetSensory(AIIntakeFormModel model, string path);
        string GetMMST(AIIntakeFormModel model, string path = "");
        string GetOtherROM(string NeckFFROM, string NeckFFLeft, string NeckFFRight, string NeckExt, string NeckExtLeft, string NeckExtRight, string path);

        // History
        string GenerateHistory(AIIntakeFormModel model);
        string GeneratePmh(string isPmh, List<string> pmhList, string pmhOther);
        string GeneratePsh(string isSmh, string pshDetails);
        string GenerateMeds(string meds, string medsDetails);
        string GenerateAllergy(string allergy, string allergyDetails);
        string GenerateGait(string gaitRadio, List<string> gaitCheckboxes, string gaitOther);

        // ROS, ADL, Occupation & Recommendations
        string GenerateRos(List<string> selectedRos, string rosOther = null);
        string GenerateRosIrina(AIIntakeFormModel model);
        string GenerateADL(AIIntakeFormModel model);
        string GenerateOccupation(string occupation, string jobTitle);
        string GenerateJob(AIIntakeFormModel model);
        string GenerateRecommandation(AIIntakeFormModel model);

        // Utility
        string CapitalizeFirstLetter(string input);
    }
}
