using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
    public class InitialIntake
    {


        public int? id { get; set; }
        public int? location_id { get; set; }
        public string? fname { get; set; }
        public string? lname { get; set; }

        public DateTime? dob { get; set; }
        public DateTime? doa { get; set; }

        public string? wcb { get; set; }
        public int? age { get; set; }
        public string? gender { get; set; }

        public string? handedness { get; set; }

        public string? rsh { get; set; }
        public string? lsh { get; set; }

        public string? rkn { get; set; }
        public string? lkn { get; set; }
        public string? lelb { get; set; }
        public string? relb { get; set; }
        public string? rhip { get; set; }
        public string? lhip { get; set; }
        public string? rank { get; set; }
        public string? lank { get; set; }
        public string? rwri { get; set; }
        public string? lwri { get; set; }

        public int? weight { get; set; }
        public int? height { get; set; }
        public string? MVA { get; set; }
        public string? workrelated { get; set; }
        public string? working_yes { get; set; }
        public string? degreeofDisability { get; set; }

        public string? asymptomaticpriortoaccident_yes { get; set; }
        public string? historyofpriortrauma_yes { get; set; }
        public string? description_of_the_priortrauma { get; set; }
        public string? Pedestrian { get; set; }

        public string? Bicyclist { get; set; }
        public string? Motorcyclist { get; set; }
        public string? Buspass { get; set; }
        public string? Driver { get; set; }
        public string? FrontPass { get; set; }
        public string? RearPass { get; set; }
        public string? Rear { get; set; }
        public string? Front { get; set; }
        public string? Driversidefront { get; set; }

        public string? Driversiderear { get; set; }
        public string? DriverSideswipe { get; set; }
        public string? TBonedDriverside { get; set; }
        public string? Passengersidefront { get; set; }
        public string? Passengersiderear { get; set; }
        public string? TBonePassengerside { get; set; }
        public string? PassengerSideswipe { get; set; }



        public string? airbagsdeployed_yes { get; set; }
        public string? EMSArrived_yes { get; set; }
        public string? PoliceatScene_yes { get; set; }
        public string? WenttoHospital_yes { get; set; }
        public string? Hospitalname { get; set; }
        public string? PMHNone { get; set; }
        public string? PMHDiabetes { get; set; }
        public string? PMHHTN { get; set; }
        public string? PMHHLD { get; set; }
        public string? PMHAsthma { get; set; }
        public string? PMHCardiac { get; set; }
        public string? PMHThyroid { get; set; }
        public string? PMHCA { get; set; }
        public string? description_of_the_PMH { get; set; }
        public string? PSHNone { get; set; }
        public string? description_of_the_PSH { get; set; }

        public string? MedsNone { get; set; }
        public string? PainmedsPRN { get; set; }
        public string? description_of_the_Meds { get; set; }
        public string? Unabletorecall { get; set; }
        public string? DrugAllergy_yes { get; set; }

        public string? DrugAllergy { get; set; }
        public string? Smoke_yes { get; set; }
        public string? ppd_txt { get; set; }
        public string? Alcohol_yes { get; set; }
        public string? Cannabis_yes { get; set; }
        public string? RecreationalDrugs_yes { get; set; }
        public string? PTChiro_yes { get; set; }
        public string? Duration_txt { get; set; }
        public string? ReliefGood { get; set; }
        public string? ReliefLittle { get; set; }
        public string? ReliefNone { get; set; }
        public string? Walk_yes { get; set; }
        public string? blocks_txt { get; set; }
        public string? Stand_yes { get; set; }
        public string? Stand_txt { get; set; }
        public string? Sit_yes { get; set; }
        public string? Sit_txt { get; set; }

        public string? Garden { get; set; }
        public string? Playsports { get; set; }
        public string? Drive { get; set; }
        public string? Lift { get; set; }
        public string? Childcare { get; set; }
        public string? Carry { get; set; }

        public string? Reachoverhead { get; set; }
        public string? Laundry { get; set; }
        public string? Shopping { get; set; }
        public string? Errands { get; set; }
        public string? Kneel { get; set; }
        public string? Squat { get; set; }
        public string? Stairs { get; set; }
        public string? Jog { get; set; }
        public string? Exercise { get; set; }
        public int cmp_id { get; set; }

        public string? JobTitle { get; set; }
        public bool? Stopped_yes { get; set; }
        public string? Retired { get; set; }
        public string? Unemployed { get; set; }
        public string? working_why { get; set; }
        public string? wr1 { get; set; }
        public string? wr2 { get; set; }
        public string? wr3 { get; set; }
        public string? wr4 { get; set; }
        public string? wr5 { get; set; }
        public string? wr6 { get; set; }
        //public string? txt_PC_RSH { get; set; }
        //public string? PC_RSH_Constant { get; set; }
        //public string? PC_RSH_Intermittent { get; set; }
        //public string? PC_RSH_Stif { get; set; }
        //public string? PC_RSH_Pop { get; set; }
        //public string? PC_RSH_Click { get; set; }

        //public string? PC_RSH_Weak { get; set; }
        //public string? RHS_Reachoverhead_yes { get; set; }
        //public string? RHS_Reachback_yes { get; set; }       

        //public string? PC_RSH_Unabletosleep { get; set; }
        //public string? PC_RSH_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_LSH { get; set; }
        //public string? PC_LSH_Constant { get; set; }
        //public string? PC_LSH_Intermittent { get; set; }
        //public string? PC_LSH_Stif { get; set; }
        //public string? PC_LSH_Weak { get; set; }
        //public string? PC_LSH_Pop { get; set; }
        //public string? PC_LSH_Click { get; set; }
        //public string? LSH_Reachoverhead_yes { get; set; }
        //public string? LSH_Reachback_yes { get; set; }
        //public string? PC_LSH_Unabletosleep { get; set; }
        //public string? PC_LSH_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_RKN { get; set; }
        //public string? PC_RKN_Constant { get; set; }
        //public string? PC_RKN_Intermittent { get; set; }
        //public string? PC_RKN_Stif { get; set; }
        //public string? PC_RKN_Weak { get; set; }
        //public string? RKN_Diffrisingfromchair_yes { get; set; }

        //public string? RKN_Diffwstairs_yes { get; set; }
        //public string? PC_RKN_Click { get; set; }
        //public string? PC_RKN_Pop { get; set; }
        //public string? PC_RKN_Buckl { get; set; }
        //public string? PC_RKN_Lock { get; set; }
        //public string? PC_RKN_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_LKN { get; set; }
        //public string? PC_LKN_Constant { get; set; }
        //public string? PC_LKN_Intermittent { get; set; }
        //public string? PC_LKN_Stif { get; set; }
        //public string? PC_LKN_Weak { get; set; }
        //public string? LKN_Diffrisingfromchair_yes { get; set; }

        //public string? LKN_Diffwstairs_yes { get; set; }
        //public string? PC_LKN_Click { get; set; }
        //public string? PC_LKN_Pop { get; set; }
        //public string? PC_LKN_Buckl { get; set; }
        //public string? PC_LKN_Lock { get; set; }
        //public string? PC_LKN_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_RHIP { get; set; }
        //public string? PC_RHIP_Constant { get; set; }
        //public string? PC_RHIP_Intermittent { get; set; }
        //public string? PC_RHIP_Lock { get; set; }
        //public string? PC_RHIP_Painwstandwalkclimb { get; set; }
        //public string? PC_RHIP_Standingfromsitting { get; set; }
        //public string? PC_RHIP_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_LHIP { get; set; }
        //public string? PC_LHIP_Constant { get; set; }
        //public string? PC_LHIP_Intermittent { get; set; }
        //public string? PC_LHIP_Lock { get; set; }
        //public string? PC_LHIP_Painwstandwalkclimb { get; set; }
        //public string? PC_LHIP_Standingfromsitting { get; set; }
        //public string? PC_LHIP_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_RANK { get; set; }
        //public string? PC_RANK_Constant { get; set; }
        //public string? PC_RANK_Intermittent { get; set; }
        //public string? PC_RANK_Painwstandwalkclimb { get; set; }
        //public string? PC_RANK_ImpwRestMedPTIce { get; set; }
        //public string? txt_PC_LANK { get; set; }
        //public string? PC_LANK_Constant { get; set; }
        //public string? PC_LANK_Intermittent { get; set; }
        //public string? PC_LANK_Painwstandwalkclimb { get; set; }
        //public string? PC_LANK_ImpwRestMedPTIce { get; set; }

        //public string? txt_PC_RWRI { get; set; }
        //public string? PC_RWRI_Constant { get; set; }
        //public string? PC_RWRI_Intermittent { get; set; }
        //public string? PC_RWRI_Weak { get; set; }
        //public string? PC_RWRI_Numb { get; set; }
        //public string? PC_RWRI_Tingle { get; set; }
        //public string? PC_RWRI_Painwliftcarrydrive { get; set; }
        //public string? PC_RWRI_ImpwRestMedPTIce { get; set; }

        //public string? txt_PC_LWRI { get; set; }
        //public string? PC_LWRI_Constant { get; set; }
        //public string? PC_LWRI_Intermittent { get; set; }
        //public string? PC_LWRI_Weak { get; set; }
        //public string? PC_LWRI_Numb { get; set; }
        //public string? PC_LWRI_Tingle { get; set; }
        //public string? PC_LWRI_Painwliftcarrydrive { get; set; }
        //public string? PC_LWRI_ImpwRestMedPTIce { get; set; }

        //public string? txt_PC_RELB { get; set; }
        //public string? PC_RELB_Constant { get; set; }
        //public string? PC_RELB_Intermittent { get; set; }
        //public string? PC_RELB_Weak { get; set; }
        //public string? PC_RELB_Numb { get; set; }
        //public string? PC_RELB_Tingle { get; set; }
        //public string? PC_RELB_Painwliftcarrydrive { get; set; }
        //public string? PC_RELB_ImpwRestMedPTIce { get; set; }

        //public string? txt_PC_LELB { get; set; }
        //public string? PC_LELB_Constant { get; set; }
        //public string? PC_LELB_Intermittent { get; set; }
        //public string? PC_LELB_Weak { get; set; }
        //public string? PC_LELB_Numb { get; set; }
        //public string? PC_LELB_Tingle { get; set; }
        //public string? PC_LELB_Painwliftcarrydrive { get; set; }
        //public string? PC_LELB_ImpwRestMedPTIce { get; set; }


        //public string? description_of_the_OtherComplaint { get; set; }

        //public string? txt_Temp { get; set; }
        //public string? txt_HR { get; set; }
        //public string? txt_RR { get; set; }
        //public string? txt_Oxygen { get; set; }
        //public string? txt_BloodPressure { get; set; }

        //public string? Fevers { get; set; }
        //public string? chills { get; set; }
        //public string? nightsweats { get; set; }
        //public string? weightgain { get; set; }
        //public string? weightloss { get; set; }
        //public string? Doublevision { get; set; }
        //public string? eyepain { get; set; }
        //public string? eyered { get; set; }
        //public string? hearingloss { get; set; }
        //public string? earache { get; set; }
        //public string? earringing { get; set; }

        //public string? nosebleeds { get; set; }

        //public string? sorethroat { get; set; }
        //public string? hoarseness { get; set; }


        //public string? Coldintolerance { get; set; }
        //public string? appetitechanges { get; set; }
        //public string? hairchanges { get; set; }
        //public string? ClearSkin { get; set; }
        //public string? norashesorlesions { get; set; }
        //public string? Headaches { get; set; }
        //public string? dizziness { get; set; }

        //public string? vertigo { get; set; }


        //public string? tremors { get; set; }

        //public string? Wheezing { get; set; }
        //public string? coughing { get; set; }
        //public string? shortnessofbreath { get; set; }
        //public string? difficultybreathing { get; set; }
        //public string? Chestpain { get; set; }
        //public string? murmurs { get; set; }
        //public string? irregularheartrate { get; set; }
        //public string? hypertension { get; set; }

        //public string? Nausea { get; set; }

        //public string? vomiting { get; set; }
        //public string? diarrhea { get; set; }
        //public string? constipation { get; set; }
        //public string? jaundice { get; set; }
        //public string? changeinbowelhabits { get; set; }
        //public string? Bloodinurine { get; set; }
        //public string? painfulurination { get; set; }
        //public string? lossofbladdercontrol { get; set; }
        //public string? urinaryretention { get; set; }
        //public string? Activebleeding { get; set; }

        //public string? bruising { get; set; }
        //public string? anemia { get; set; }

        //public string? bloodclottingdisorders { get; set; }
        //public string? Anxiety { get; set; }
        //public string? changeinsleeppattern { get; set; }
        //public string? depression { get; set; }
        //public string? suicidalthoughts { get; set; }
        //public string? plan_Imaging { get; set; }
        //public string? plan_Alltreatment { get; set; }
        //public string? plan_ColdCompresses { get; set; }

        //public string? plan_Continueantiinflammatory { get; set; }
        //public string? plan_Continuewithphysicaltherapy { get; set; }

        //public string? plan_Wewillreevaluate { get; set; }

        //public string? plan_steroidinjections { get; set; }
        //public string? plan_Historyofinjury { get; set; }
        //public string? txt_plan_Historyofinjury1 { get; set; }
        //public string? txt_plan_Historyofinjury2 { get; set; }
        //public string? txt_plan_Historyofinjury3 { get; set; }
        //public string? plan_recommendoperative { get; set; }
        //public string? plan_medicallynecessary { get; set; }
        //public string? plan_Discussedthelength { get; set; }
        //public string? plan_Allthebenefits { get; set; }
        //public string? plan_questions { get; set; }

        //public string? plan_verbally { get; set; }
        //public string? txt_verbally { get; set; }
        //public string? plan_mc { get; set; }

        //public string? plan_followup12weekspostop { get; set; }



        //public string? plan_Followupin2weeks { get; set; }
        //public string? plan_Referral { get; set; }
        //public string? plan_injforpainmgmt { get; set; }
        //public string? inj_RSH { get; set; }
        //public string? inj_LSH { get; set; }

        //public string? inj_RKN { get; set; }


        //public string? inj_LKN { get; set; }
        //public string? inj_RHIP { get; set; }
        //public string? inj_LHIP { get; set; }
        //public string? inj_RANK { get; set; }
        //public string? inj_LANK { get; set; }
        //public string? inj_RWRI { get; set; }

        //public string? inj_LWRI { get; set; }
        //public string? inj_RELB { get; set; }
        //public string? inj_LELB { get; set; }
        //public string? plan_MRIordered { get; set; }
        //public string? MRIordered_RSH { get; set; }

        //public string? MRIordered_LSH { get; set; }
        //public string? MRIordered_RKN { get; set; }
        //public string? MRIordered_LKN { get; set; }
        //public string? MRIordered_RHIP { get; set; }
        //public string? MRIordered_LHIP { get; set; }
        //public string? MRIordered_RANK { get; set; }
        //public string? MRIordered_LANK { get; set; }
        //public string? MRIordered_RWRI { get; set; }
        //public string? MRIordered_LWRI { get; set; }
        //public string? MRIordered_RELB { get; set; }
        //public string? MRIordered_LELB { get; set; }
        //public string? conservativemanagement { get; set; }
        //public string? fuweeks { get; set; }
        //public string? txt_plan_fuweeks { get; set; }

        //public string? plan_recommend { get; set; }
        //public string? plan_recommend_RSH { get; set; }
        //public string? plan_recommend_LSH { get; set; }
        //public string? plan_recommend_RKN { get; set; }
        //public string? plan_recommend_LKN { get; set; }
        //public string? plan_recommend_RHIP { get; set; }

        //public string? plan_recommend_LHIP { get; set; }
        //public string? plan_recommend_RANK { get; set; }
        //public string? plan_recommend_LANK { get; set; }
        //public string? plan_recommend_RWRI { get; set; }
        //public string? plan_recommend_LWRI { get; set; }

        //public string? plan_recommend_RELB { get; set; }
        //public string? plan_recommend_LELB { get; set; }
        //public string? ProceedwSx { get; set; }
        //public string? Wantstothinkaboutit { get; set; }
        //public string? MedClearance { get; set; }
        //public string? WC_authorization { get; set; }
        //public string? Patientconsentsto { get; set; }
        //public string? txt_Patientconsentsto { get; set; }

        //public string? plan_Patientscheduledfor { get; set; }


        //public string? plan_Patientscheduledfor_recommend_RSH { get; set; }

        //public string? plan_Patientscheduledfor_recommend_LSH { get; set; }
        //public string? plan_Patientscheduledfor_recommend_RKN { get; set; }
        //public string? plan_Patientscheduledfor_recommend_LKN { get; set; }
        //public string? plan_Patientscheduledfor_recommend_RHIP { get; set; }
        //public string? plan_Patientscheduledfor_recommend_LHIP { get; set; }
        //public string? plan_Patientscheduledfor_recommend_RANK { get; set; }
        //public string? plan_Patientscheduledfor_recommend_LANK { get; set; }
        //public string? plan_Patientscheduledfor_recommend_RWRI { get; set; }
        //public string? plan_Patientscheduledfor_recommend_LWRI { get; set; }
        //public string? plan_Patientscheduledfor_recommend_RELB { get; set; }
        //public string? plan_Patientscheduledfor_recommend_LELB { get; set; }



        public int? Cmp_Id { get; set; }

    }

    public class IntakeDropDown
    {
        public string Name { get; set; }
    }

    public class InitialIntakeAI
    {
        public int Id { get; set; }
        public int? LocationId { get; set; }
        public string? Location { get; set; }
        public string FN { get; set; }
        public string LN { get; set; }
        public string Email { get; set; }
        public string Mobile { get; set; }
        public string DLPath { get; set; }

        public DateTime? DOB { get; set; }

        public DateTime? DOE { get; set; }

        public DateTime? DOA { get; set; }
        public DateTime? SendDate { get; set; }
        public DateTime? PatientSubmitDate { get; set; }

        public string Visit_Type { get; set; }
        public string Diagnosis { get; set; }

        public string TreatmentIds { get; set; }
        public string TreatmentDelimitDesc { get; set; }
        public string TreatmentDesc { get; set; }

        // JSON column → store as string OR strongly typed object
        public string FormData { get; set; }

        public int? CmpId { get; set; }
    }
}
