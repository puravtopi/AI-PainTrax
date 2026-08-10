using PainTrax.Web.ViewModel;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Helper
{
    public static class ReportHelper
    {


        public static string GetDisabilityStatement(AIIntakeFormModel data)
        {
            string pct = data.degree ?? "";

            string isNotWorking = (data.IsWorking == "Yes") ? "working" : "not working";
            string isPartially = (data.YesWork == "Parttime") ? "part-time" : "full-time";

            if (isNotWorking == "not working")
                isPartially = "";

            return $"{pct}% - temporarily partially disabled - The patient is currently {isNotWorking} {isPartially}.";
        }

        public static string GetHistory(AIIntakeFormModel model)
        {
            // 1. Basic Info
            string genderStr = model.Gender?.ToLower() == "male" ? "male" : "female";
            string handedStr = model.DominantHand == "right-handed" ? "right-handed" : "left-handed";
            string restrainedStr = model.SeatBelt == "Yes" ? "restrained" : "unrestrained";

            // 2. Accident Type Logic
            string accidentTypeStr = (model.InjuryType == "WC") ? "work-related accident" : "motor vehicle accident";
            string doaStr = model.DOA ?? "____";

            // 3. Collision Type (Join array)
            string collisionStr = model.PatientAccidentType != null ? string.Join("/", model.PatientAccidentType) : "____";

            // 4. Airbags and EMS
            string airbagStr = model.Airbagsdeployed == "Yes" ? "deployed" : "did not deploy";
            string emsStr = model.EMS == "Yes" ? "arrived" : "did not arrive";

            // 5. LOC and Bruises
            string locBruiseNarrative = "";
            if (model.LOC == "Yes")
            {
                locBruiseNarrative = $"The patient reports loss of consciousness ({model.OtherLOC}) and bruises ({model.OtherBruises ?? " noted"}).";
            }
            else
            {
                locBruiseNarrative = "The patient denies loss of consciousness.";
            }

            // 6. Complaints (Skeletal injuries)
            var allComplaints = new List<string>();
            if (model.Complaints != null) allComplaints.AddRange(model.Complaints);
            if (!string.IsNullOrEmpty(model.OtherBodyPart)) allComplaints.Add(model.OtherBodyPart);
            string injuriesStr = allComplaints.Count > 0 ? string.Join(", ", allComplaints) : "________";

            // 7. Hospital Logic
            string hospitalNarrative = "";
            if (model.Hospital == "Yes")
            {
                hospitalNarrative = $"The patient was taken on an emergent basis to {model.HospitalName} Hospital where the patient was treated and released.";
            }
            else
            {
                hospitalNarrative = "The patient did not go to any hospital that same day.";
            }

            // Final String Assembly
            return $"The patient is a {model.Age}-year-old {handedStr} {genderStr}, " +
                   $"who was the {restrainedStr} {string.Join(", ", model.PatientAccidentType)} of a vehicle that was involved in a {accidentTypeStr} on {doaStr}. " +
                   $"The patient's vehicle was impacted on the {collisionStr}. " +
                   $"The airbags {airbagStr}. The EMS {emsStr} on the scene. " +
                   $"{locBruiseNarrative} The patient sustained multiple skeletal injuries including injury to {injuriesStr}. " +
                   $"{hospitalNarrative} The patient has been undergoing physical therapy for the past ___ weeks/months. " +
                   $"My evaluation is limited to {injuriesStr} injury sustained in the accident of {doaStr}.";
        }

        public static string GetDiagnosis(string html)
        {
            // 1. Extract the text inside each <li> tag
            var matches = Regex.Matches(html, @"<li>(.*?)</li>", RegexOptions.IgnoreCase);

            // 2. Add the number and a space to each item
            var lines = matches.Cast<Match>()
                .Select((m, index) => $"{index + 1}. {m.Groups[1].Value.Trim()}");

            // 3. Join them using a NewLine character so each starts on a new line
            return string.Join(Environment.NewLine, lines);
        }

        public static string GetPalpation(dynamic model)
        {
            // 1. Tenderness Logic
            // Handles: No/mild/moderate/severe ... bilaterally/right/left
            string palpationSeverity = model.Palpation ?? "No";
            string palpationSide = model.Palpationtenderness ?? "";

            string tendernessPart = $"{palpationSeverity} tenderness to palpation of the midline and paraspinal muscles {palpationSide}".Trim();

            // 2. Trigger Points Logic (Checkboxes)
            // Joins multiple selected items with a comma
            string triggerPointsText = "none";
            if (model.PETriggerCervical != null && model.PETriggerCervical.Count > 0)
            {
                triggerPointsText = string.Join(", ", model.PETriggerCervical);
            }

            string triggerPart = $"Trigger points noted in {triggerPointsText}";

            // 3. Muscle Spasm Logic
            // Handles: none/mild/moderate/severe ... bilaterally/right/left
            string spasmSeverity = model.PEMuscleCervical ?? "none";
            string spasmSide = (spasmSeverity == "none") ? "" : (model.PEMuscleCervicalSide ?? "");

            string spasmPart = $"Muscle spasms are {spasmSeverity} {spasmSide}".Trim();

            // Combine all parts into the final paragraph
            return $"{tendernessPart}. {triggerPart}. {spasmPart}.";
        }

    }
}
