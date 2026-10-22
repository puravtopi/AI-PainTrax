using System.Text.Json.Serialization;

namespace PainTrax.Web.Models
{
    public class DetailedClinicalReport
    {
        [JsonPropertyName("PatientInformation")]
        public string PatientInformation { get; set; } = string.Empty;

        [JsonPropertyName("WorkAndAdministrativeDetails")]
        public string WorkAndAdministrativeDetails { get; set; } = string.Empty;

        [JsonPropertyName("ProceduresPerformed")]
        public string ProceduresPerformed { get; set; } = string.Empty;

        [JsonPropertyName("ActivitiesOfDailyLiving")]
        public string ActivitiesOfDailyLiving { get; set; } = string.Empty;

        [JsonPropertyName("PresentComplaints")]
        public string PresentComplaints { get; set; } = string.Empty;

        [JsonPropertyName("History")]
        public string History { get; set; } = string.Empty;

        [JsonPropertyName("ReviewOfSystems")]
        public string ReviewOfSystems { get; set; } = string.Empty;

        [JsonPropertyName("PastMedicalHistory")]
        public string PastMedicalHistory { get; set; } = string.Empty;

        [JsonPropertyName("PastSurgicalHistory")]
        public string PastSurgicalHistory { get; set; } = string.Empty;

        [JsonPropertyName("SocialHistory")]
        public string SocialHistory { get; set; } = string.Empty;

        [JsonPropertyName("Medications")]
        public string Medications { get; set; } = string.Empty;

        [JsonPropertyName("Allergies")]
        public string Allergies { get; set; } = string.Empty;

        [JsonPropertyName("Occupation")]
        public string Occupation { get; set; } = string.Empty;

        [JsonPropertyName("VitalSigns")]
        public string VitalSigns { get; set; } = string.Empty;

        [JsonPropertyName("NeurologicalExamination")]
        public string NeurologicalExamination { get; set; } = string.Empty;

        [JsonPropertyName("CoordinationExamination")]
        public string CoordinationExamination { get; set; } = string.Empty;

        [JsonPropertyName("ReflexExamination")]
        public string ReflexExamination { get; set; } = string.Empty;

        [JsonPropertyName("SensoryExamination")]
        public string SensoryExamination { get; set; } = string.Empty;

        [JsonPropertyName("MotorExamination")]
        public string MotorExamination { get; set; } = string.Empty;

        [JsonPropertyName("MusculoskeletalExamination")]
        public string MusculoskeletalExamination { get; set; } = string.Empty;

        [JsonPropertyName("RangeOfMotion")]
        public string RangeOfMotion { get; set; } = string.Empty;

        [JsonPropertyName("Gait")]
        public string Gait { get; set; } = string.Empty;

        [JsonPropertyName("DiagnosticStudies")]
        public string DiagnosticStudies { get; set; } = string.Empty;

        [JsonPropertyName("Diagnoses")]
        public string Diagnoses { get; set; } = string.Empty;

        [JsonPropertyName("Recommendations")]
        public string Recommendations { get; set; } = string.Empty;

        [JsonPropertyName("CareRestrictions")]
        public string CareRestrictions { get; set; } = string.Empty;

        [JsonPropertyName("Precautions")]
        public string Precautions { get; set; } = string.Empty;

        [JsonPropertyName("Goals")]
        public string Goals { get; set; } = string.Empty;

        [JsonPropertyName("FollowUp")]
        public string FollowUp { get; set; } = string.Empty;

        [JsonPropertyName("Causation")]
        public string Causation { get; set; } = string.Empty;

        // Additional property to hold the raw transcript for the frontend response
        [JsonIgnore]
        public string? Transcript { get; set; }
    }
}
