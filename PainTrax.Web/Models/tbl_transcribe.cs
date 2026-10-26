namespace PainTrax.Web.Models
{
    public class tbl_transcribe
    {
        public int id { get; set; }
        public string content { get; set; }
        public string history { get; set; }
        public string cc { get; set; }
        public string plan { get; set; }
        public string subjective { get; set; }
        public string assessment { get; set; }
        public string file_path { get; set; }
        public long ie_id { get; set; }
        public long fu_id { get; set; }
        public int length { get; set; }
        public int intake_id { get; set; }
        public int cmp_id { get; set; }
        public DateTime date { get; set; }

        public string PatientInformation { get; set; } = string.Empty;

        public string WorkAndAdministrativeDetails { get; set; } = string.Empty;

        public string ProceduresPerformed { get; set; } = string.Empty;

        public string ActivitiesOfDailyLiving { get; set; } = string.Empty;

        public string ReviewOfSystems { get; set; } = string.Empty;

        public string PastMedicalHistory { get; set; } = string.Empty;

        public string PastSurgicalHistory { get; set; } = string.Empty;

        public string SocialHistory { get; set; } = string.Empty;

        public string Medications { get; set; } = string.Empty;

        public string Allergies { get; set; } = string.Empty;

        public string Occupation { get; set; } = string.Empty;

        public string VitalSigns { get; set; } = string.Empty;

        public string NeurologicalExamination { get; set; } = string.Empty;

        public string CoordinationExamination { get; set; } = string.Empty;

        public string ReflexExamination { get; set; } = string.Empty;

        public string SensoryExamination { get; set; } = string.Empty;

        public string MotorExamination { get; set; } = string.Empty;

        public string MusculoskeletalExamination { get; set; } = string.Empty;

        public string RangeOfMotion { get; set; } = string.Empty;

        public string Gait { get; set; } = string.Empty;

        public string DiagnosticStudies { get; set; } = string.Empty;

        public string Diagnoses { get; set; } = string.Empty;

        public string Recommendations { get; set; } = string.Empty;

        public string CareRestrictions { get; set; } = string.Empty;

        public string Precautions { get; set; } = string.Empty;

        public string Goals { get; set; } = string.Empty;

        public string FollowUp { get; set; } = string.Empty;

        public string Causation { get; set; } = string.Empty;

    }
}
