namespace PainTrax.Web.Models
{
    public class PatientDetails
    {
        public string? id { get; set; }
        public string? signatureData { get; set; }
        public string? fname { get; set; }
        public string? lname { get; set; }
        public DateTime? dob { get; set; }
        public bool? isExist { get; set; }
        public DateTime? executionDate { get; set; }
        public string? primaryPhone { get; set; }
        public string? alternatePhone { get; set; }
        public string? email { get; set; }
        public string? patientPortalUsername { get; set; }

        //// Section 1 - Electronic Communication

        public string? sMSConsent { get; set; } // yes / no
        public string? sMSYesInitials { get; set; }
        public string? sMSNoInitials { get; set; }

        public string? emailConsent { get; set; }
        public string? emailYesInitials { get; set; }
        public string? emailNoInitials { get; set; }

        public string? voiceMailConsent { get; set; }
        public string? voiceMailYesInitials { get; set; }
        public string? voiceMailNoInitials { get; set; }

        public string? patientPortalConsent { get; set; }
        public string? patientPortalYesInitials { get; set; }
        public string? patientPortalNoInitials { get; set; }

        // Section 2 - Audio Recording

        public string? audioRecordingConsent { get; set; } // yes / no
        public string? audioYesInitials { get; set; }
        public string? audioNoInitials { get; set; }

        // Section 3 - Marketing Authorization

        public string? marketingConsent { get; set; } // yes / no
        public string? marketingYesInitials { get; set; }
        public string? marketingNoInitials { get; set; }

        //// Signature Section

        public string? signedByName { get; set; }
        public DateTime? signedDate { get; set; }
        public string? relationshipToPatient { get; set; }

        public string? patientId { get; set; }
    }
}
