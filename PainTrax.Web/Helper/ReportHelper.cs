using PainTrax.Web.ViewModel;
using System.Text;
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
                hospitalNarrative = $"The patient was taken on an emergent basis to {model.hospitalname} Hospital where the patient was treated and released.";
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

        public static string GenerateCervicalSpineReport(AIIntakeFormModel model)
        {
            var report = new StringBuilder("CERVICAL SPINE: ");

            // 1. Pain and Stiffness
            string stiffnessText = model.NeckStiffness.Contains("Stiffness") ? "pain and stiffness" : "pain";
            report.Append($"The patient complains of {stiffnessText} in the neck region. ");

            // 2. Radiation
            if (!string.IsNullOrEmpty(model.NeckRadiatesTo) || model.NeckRadiates.Any())
            {
                string side = model.NeckRadiatesTo ?? "";
                if (side == "Bilateral") side = "bilateral";

                string radiationAreas = model.NeckRadiates.Any()
                    ? string.Join(", ", model.NeckRadiates).ToLower()
                    : "areas";

                report.Append($"The pain radiates from the neck to {side} {radiationAreas} ");
            }

            // 3. Associated Symptoms
            if (model.NeckAssociated.Any())
            {
                string associated = string.Join(", ", model.NeckAssociated).ToLower();
                // Replace last comma with 'and' for better grammar
                int lastComma = associated.LastIndexOf(',');
                if (lastComma != -1)
                    associated = associated.Remove(lastComma, 1).Insert(lastComma, " and");

                report.Append($"and associated with {associated}. ");
            }

            // 4. Difficulties (Turning/Gripping)
            var difficulties = new List<string>();
            if (model.NeckStiffness.Contains("Diff turning") || model.NeckStiffness.Contains("rotating head"))
                difficulties.Add("turning/rotating the head");
            if (model.NeckStiffness.Contains("Diff gripping hand"))
                difficulties.Add("gripping objects");

            if (difficulties.Any())
            {
                report.Append($"The patient has difficulty {string.Join(" and ", difficulties)}. ");
            }

            // 5. PT Improvement
            if (!string.IsNullOrEmpty(model.NeckSustainedStiffness))
            {
                report.Append($"{model.NeckSustainedStiffness}. ");
            }

            // 6. Worsening/Improving Factors (Optional based on your UI)
            if (model.NeckWorsens.Any())
                report.Append($"Symptoms are worsened by {string.Join(", ", model.NeckWorsens).ToLower()}. ");

            // 7. Pain Score
            report.Append($"Pain score is {model.NeckPain ?? "0"}/10.");

            return report.ToString();
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
