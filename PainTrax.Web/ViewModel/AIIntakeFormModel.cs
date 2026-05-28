namespace PainTrax.Web.ViewModel
{
    public class AIIntakeFormModel
    {
        /* public string Id { get; set; }
         public string FN { get; set; }
         public string LN { get; set; }
         public string Gender { get; set; }
         public string DominantHand { get; set; }
         public string DOB { get; set; }
         public string DOA { get; set; }
         public string DOE { get; set; }
         public int Age { get; set; }
         public string Height { get; set; }
         public string Weight { get; set; }
         public string IsWorking { get; set; }
         public string JobTitle { get; set; }
         public string StoppedAfterAccident { get; set; }
         public string NotWorkingReason { get; set; }
         public List<string> Complaints { get; set; }
         public string InjuryType { get; set; }
         public string Activity { get; set; }
         public string Incident { get; set; }
         public string IncidentType { get; set; }
         public string Mechanism { get; set; }
         public string SymptomOnset { get; set; }
         public string BodyPart { get; set; }
         public string SymptomPattern { get; set; }
         public string DailyActivities { get; set; }*/

        public string Id { get; set; }
        public string FN { get; set; }
        public string LN { get; set; }
        public string Age { get; set; }
        public string LocationId { get; set; }
        public string DOA { get; set; }
        public string DOB { get; set; }
        public string DOE { get; set; }
        public string EMS { get; set; }
        public List<string> PMH { get; set; }
        public string PPD { get; set; }
        public string PSH { get; set; }
        public string Sit { get; set; }
        public string Meds { get; set; }
        public string Walk { get; set; }
        public string HdId { get; set; }
        public string Chiro { get; set; }
        public string Smoke { get; set; }
        public string Stand { get; set; }
        public string Airbag { get; set; }
        public string Gender { get; set; }
        public string Height { get; set; }
        public string Heightft { get; set; }
        public string Heightinch { get; set; }
        public string PTFreq { get; set; }
        public string Police { get; set; }
        public string SitMin { get; set; }
        public List<string> Unable { get; set; }
        public string Weight { get; set; }
        public string LblDOA { get; set; }
        public string Alcohol { get; set; }
        public string Allergy { get; set; }
        public string LKnPain { get; set; }
        public List<string> LKnType { get; set; }
        public string LShPain { get; set; }
        public List<string> LShType { get; set; }
        public string RKnPain { get; set; }
        public List<string> RKnType { get; set; }
        public List<string> RShInsp { get; set; }
        public string RShPain { get; set; }
        public string RShReachOverhead { get; set; }
        public string RShReachBack { get; set; }
        public string RShSleepIssue { get; set; }
        public List<string> RShPalp { get; set; }
        public string RShProm { get; set; }
        public List<string> RShType { get; set; }
        public string Records { get; set; }
        public string Activity { get; set; }
        public string BodyPart { get; set; }
        public string Cannabis { get; set; }
        public string Hospital { get; set; }
        public string Incident { get; set; }
        public string JobTitle { get; set; }
        public string PMHOther { get; set; }
        public string PTRelief { get; set; }
        public List<string> RShNeuro { get; set; }
        public List<string> RShTests { get; set; }
        public string StandMin { get; set; }
        public string ChiroFreq { get; set; }
        public string IsWorking { get; set; }
        public string Mechanism { get; set; }
        public string MedsOther { get; set; }
        public string PreInjury { get; set; }
        public string Transport { get; set; }
        public List<string> Complaints { get; set; }
        public string Compliance { get; set; }
        public string InjuryType { get; set; }
        public List<string> LKnImprove { get; set; }
        public List<string> LShImprove { get; set; }
        public string OtherDrugs { get; set; }
        public string PTDuration { get; set; }
        public string PainEffect { get; set; }
        public List<string> RKnImprove { get; set; }
        public string RShAbdPain { get; set; }

        public string RShExtPain { get; set; }
        public List<string> RShImprove { get; set; }
        public string RshFlexion { get; set; }
        public List<string> VehicleHit { get; set; }
        public string WalkBlocks { get; set; }
        public string AllergyDrug { get; set; }
        public string ChiroRelief { get; set; }
        public List<string> LKnSymptoms { get; set; }
        public List<string> LShSymptoms { get; set; }
        public string LKnReachOverhead { get; set; }
        public string LKnReachBack { get; set; }
        public string LKnSleepIssue { get; set; }
        public List<string> PatientType { get; set; }
        public List<string> RKnSymptoms { get; set; }
        public string RShExternal { get; set; }
        public string RKnReachOverhead { get; set; }
        public string RKnReachBack { get; set; }
        public string RKnSleepIssue { get; set; }
        public string RShFlexPain { get; set; }
        public string RShInternal { get; set; }
        public List<string> RShRomLimit { get; set; }
        public List<string> RShStrength { get; set; }
        public List<string> RShSymptoms { get; set; }
        public string AccidentType { get; set; }
        public string DominantHand { get; set; }
        public string HospitalName { get; set; }
        public string IncidentType { get; set; }
        public string PatientFName { get; set; }
        public string PatientLName { get; set; }
        public string PoliceReport { get; set; }
        public string RShAbduction { get; set; }
        public string SymptomOnset { get; set; }
        public string ChiroDuration { get; set; }
        public string LShSleepIssue { get; set; }
        public string OtherDrugsTxt { get; set; }
        public string SymptomsStart { get; set; }
        public bool TreatmentPlan { get; set; }
        public string SymptomPattern { get; set; }
        public string AllergyReaction { get; set; }
        public string DailyActivities { get; set; }
        public string MechanismChoice { get; set; }
        public string LShReachOverhead { get; set; }
        public string LShReachBack { get; set; }
        public string NotWorkingReason { get; set; }
        public string PreInjuryDetails { get; set; }
        public string SymptomsProgress { get; set; }
        public string DiscontinueReason { get; set; }
        public string MechanismInvolving { get; set; }
        public string MechanismResulting { get; set; }
        public string AccidentDescription { get; set; }
        public string StoppedAfterAccident { get; set; }
        public string __RequestVerificationToken { get; set; }
        public string PatientSubmitDate { get; set; }
        public string DLPath { get; set; }
        public string Diagnosis { get; set; }
        public string TreatmentIds { get; set; }
        public string TreatmentDesc { get; set; }
        public string TreatmentDelimitDesc { get; set; }
        public string Treatment { get; set; }
        public string AccidentAudio { get; set; }


        // Saranya
        public string Occupation { get; set; }
        public string YesWork { get; set; }
        public string degree { get; set; }
        public string Asymptomatic { get; set; }
        public string Priortrauma { get; set; }
        public string PatientAT { get; set; }
        public List<string> PatientAccidentType { get; set; }
        public string SeatBelt { get; set; }
        public string Airbagsdeployed { get; set; }
        public string LOC { get; set; }
        public string Bruises { get; set; }
        public string privatecar { get; set; }
        public string OtherBodyPart { get; set; }
        public string PT { get; set; }
        public string PProcedures { get; set; }
        public string NeckPain { get; set; }
        public List<string> NeckStiffness { get; set; }
        //public string NeckDiffturning { get; set; }
        //public string Neckrotatinghead { get; set; }
        //public string NeckDiffgrippinghand { get; set; }
        //public string Necksustained { get; set; }
        public string NeckRadiatesTo { get; set; }
        public List<string> NeckRadiates { get; set; }
        public List<string> NeckAssociated { get; set; }
        public List<string> NeckWorsens { get; set; }
        public List<string> NeckImproves { get; set; }

        public string MidbackSection { get; set; }
        public List<string> mdbackPain { get; set; }
        //public string mdbackDiffsleeping { get; set; }
        //public string mdbackDifflifting { get; set; }
        //public string mdbackDiffbending { get; set; }
        //public string mdbacksustained { get; set; }
        public string MidbackRadiatesTo { get; set; }
        public List<string> MidbackRadiates { get; set; }
        public List<string> MidbackWorsens { get; set; }
        public List<string> MidbackImproves { get; set; }

        public string LowBackPain { get; set; }
        public List<string> lbackPain { get; set; }
        //public string lbackDiffsleeping { get; set; }
        //public string lbackDifflifting { get; set; }
        //public string lbackDiffbending { get; set; }
        //public string lbacksustained { get; set; }
        public string LowBackRadiatesTo { get; set; }
        public List<string> LowBackRadiates { get; set; }
        public List<string> LowBackAssociated { get; set; }
        public List<string> LowBackWorsens { get; set; }
        public List<string> LowBackImproves { get; set; }

        //public string GeneralNormal { get; set; }
        public string GeneralROS { get; set; }
        //public string SkinNormal { get; set; }
        public string SkinROS { get; set; }
       // public string HEENTNormal { get; set; }
        public string HEENTROS { get; set; }

        //public string NeckNormal { get; set; }
        public string NeckROS { get; set; }
        //public string CardiovascularNormal { get; set; }
        public string CardiovascularROS { get; set; }
        //public string RespiratoryNormal { get; set; }
        public string RespiratoryROS { get; set; }

       // public string GastrointestinalNormal { get; set; }
        public string GastrointestinalROS { get; set; }
       // public string UrinaryNormal { get; set; }
        public string UrinaryROS { get; set; }
        //public string PeripheralvascularNormal { get; set; }
        public string PeripheralvascularROS { get; set; }

        //public string MusculoskeletalNormal { get; set; }
        public string MusculoskeletalROS { get; set; }
        //public string NeurologicalNormal { get; set; }
        public string NeurologicalROS { get; set; }
       // public string EndocrineNormal { get; set; }
        public string EndocrineROS { get; set; }

        public List<string> CervicalPE { get; set; }
        public string Palpation { get; set; }
        public string Palpationtenderness { get; set; }
        public List<string> PETriggerCervical { get; set; }
        public string PEMuscleCervical { get; set; }
        public string PEMuscleCervicalSide { get; set; }

        public string CervicalFlexion { get; set; }
        public string CervicalExtension { get; set; }
        public string CervicalLeftRotation { get; set; }
        public string CervicalRightRotation { get; set; }
        public string CervicalLeftLateralbending { get; set; }
        public string CervicalrightLateralbending { get; set; }

        public string CERVICALRightDeltoid { get; set; }
        public string CERVICALRightBiceps { get; set; }
        public string CERVICALRightTriceps { get; set; }
        public string CERVICALRightWristext { get; set; }
        public string CERVICALRightWristflex { get; set; }
        public string CERVICALRightIntrinsic { get; set; }
        public string CERVICALLEFTDeltoid { get; set; }
        public string CERVICALLEFTBiceps { get; set; }
        public string CERVICALLEFTTriceps { get; set; }
        public string CERVICALLeftWristext { get; set; }
        public string CERVICALLEFTWristflex { get; set; }
        public string CERVICALLEFTIntrinsic { get; set; }

        public string CERVICALRightC5 { get; set; }
        public string CERVICALRightC6 { get; set; }
        public string CERVICALRightC7 { get; set; }
        public string CERVICALRightC8 { get; set; }
        public string CERVICALRightT1 { get; set; }

        public string CERVICALLeftC5 { get; set; }
        public string CERVICALLeftC6 { get; set; }
        public string CERVICALLeftC7 { get; set; }
        public string CERVICALLeftC8 { get; set; }
        public string CERVICALLeftT1 { get; set; }

        public string CevicalSpurling { get; set; }
        public List<string> CevicalSpurlingRight { get; set; }
        public string CERVICALRighttxt { get; set; }
        public string CERVICALLefttxt { get; set; }
        public string CERVICALbilateraltxt { get; set; }

        public string CevicalCompression { get; set; }
        public string CevicalCompressionRight { get; set; }

        public string THORACICBicepstendonright { get; set; }
        public string THORACICBicepstendonleft { get; set; }

        public string THORACICTricepstendonright { get; set; }
        public string THORACICTricepstendonleft { get; set; }
        public string CERVICALSPINEEXAMtxt { get; set; }
        public string THORACICEXAMtxt { get; set; }
        public string LUMBAREXAMtxt { get; set; }
        public List<string> THORACICPE { get; set; }
        public string THORACICPalpation { get; set; }
        public string THORACICPalpationmidline { get; set; }
        public List<string> PETriggerTHORACIC { get; set; }
        public string PEMuscleTHORACIC { get; set; }

        public string THORACICFlexion { get; set; }
        public string THORACICExtension { get; set; }
        public string THORACICLeftRotation { get; set; }
        public string THORACICRightRotation { get; set; }
        public string THORACICLeftLateralbending { get; set; }
        public string THORACICrightLateralbending { get; set; }

        public List<string> LUMBARPE { get; set; }
        public string LUMBARPalpation { get; set; }
        public string LUMBARPalpationmidline { get; set; }
        public List<string> PETriggerLUMBAR { get; set; }
        public string PEMuscleLUMBAR { get; set; }
        public string PEMuscleLUMBARside { get; set; }

        public string LUMBARFlexion { get; set; }
        public string LUMBARExtension { get; set; }
        public string LUMBARLeftRotation { get; set; }
        public string LUMBARRightRotation { get; set; }
        public string LUMBARLeftLateralbending { get; set; }
        public string LUMBARightLateralbending { get; set; }

        public string LUMBARRightIliopsoas { get; set; }
        public string LUMBARRightQuadriceps { get; set; }
        public string LUMBARRightHamstrings { get; set; }
        public string LUMBARRightTibant { get; set; }
        public string LUMBARRightEHL { get; set; }
        public string LUMBARRightGS { get; set; }

        public string LUMBARLEFTIliopsoas { get; set; }
        public string LUMBARLEFTQuadriceps { get; set; }
        public string LUMBARLEFTHamstrings { get; set; }
        public string LUMBARLEFTTibant { get; set; }
        public string LUMBARLEFTEHL { get; set; }
        public string LUMBARLEFTGS { get; set; }

        public string LUMBARRightL2 { get; set; }
        public string LUMBARRightL3 { get; set; }
        public string LUMBARRightL4 { get; set; }
        public string LUMBARRightL5 { get; set; }
        public string LUMBARRightS1 { get; set; }

        public string LUMBARLeftL2 { get; set; }
        public string LUMBARLeftL3 { get; set; }
        public string LUMBARLeftL4 { get; set; }
        public string LUMBARLeftL5 { get; set; }
        public string LUMBARLeftS1 { get; set; }

        public string LumbarStraight { get; set; }
        public List<string> LumbarStraightRight { get; set; }
        public string LumbarStraightRighttxt { get; set; }
        public string LumbarStraightLefttxt { get; set; }
        public string LumbarStraightbilateraltxt { get; set; }
        
        public string LumbarFacetloading { get; set; }
        public string LumbarFacetloadingRight { get; set; }

        public string LUMBARPatellartendonright { get; set; }
        public string LUMBARPatellartendonleft { get; set; }

        public string LUMBARAchillestendonright { get; set; }
        public string LUMBARAchillestendonleft { get; set; }

        public List<string> GAIT { get; set; }

        //public List<string> PlanConservativeTx { get; set; }
        public string PlanStart { get; set; }

        public List<string> PlanPT { get; set; }
        //public List<string> MedicationRx { get; set; }
        public List<string> PlanMedication { get; set; }
        //public List<string> PlanOrderMRI { get; set; }
        public List<string> PlanMRI { get; set; }
        public List<string> PlanCT { get; set; }
        public List<string> PlanXray { get; set; }
        public string PlanOrder { get; set; }
        //public List<string> PlanOrderEMG { get; set; }
        public List<string> PlanEMG { get; set; }
        public List<string> PlanNCV { get; set; }
        public List<string> PlanUTPI { get; set; }
        public List<string> PlanImaging { get; set; }
        public List<string> Plantreatment { get; set; }
        //public List<string> PlanRecommendation { get; set; }
        public List<string> Recommendation { get; set; }
        public string FollowUpOther { get; set; }
        public string FollowUp { get; set; }
        public string OtherAsymptomatic { get; set; }
        public string OtherPriortrauma { get; set; }
        public string OtherLOC { get; set; }
        public string OtherBruises { get; set; }

        public string PatientIEId { get; set; }
        public string PatientId { get; set; }
    }
}

