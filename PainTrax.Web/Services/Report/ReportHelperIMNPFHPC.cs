using PainTrax.Web.ViewModel;
using System.Text;
using System.Text.RegularExpressions;
namespace PainTrax.Web.Services.Report
{
    public class ReportHelperIMNPFHPC : IReportHelper
    {
        public string GetDisabilityStatement(AIIntakeFormModel data)
        {
            string pct = data.degree ?? "";

            string isNotWorking = (data.IsWorking == "Yes") ? "working" : "not working";
            string isPartially = (data.YesWork == "Parttime") ? "part-time" : "full-time";

            if (isNotWorking == "not working")
                isPartially = "";

            return $"{pct}% - temporarily partially disabled - The patient is currently {isNotWorking} {isPartially}.";
        }

        public string GetHistory(AIIntakeFormModel model)
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

        public string GenerateCervicalSpineReport(AIIntakeFormModel model)
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

        public string GetDiagnosis(string html)
        {
            // 1. Extract the text inside each <li> tag
            var matches = Regex.Matches(html, @"<li>(.*?)</li>", RegexOptions.IgnoreCase);

            // 2. Add the number and a space to each item
            var lines = matches.Cast<Match>()
                .Select((m, index) => $"{index + 1}. {m.Groups[1].Value.Trim()}");

            // 3. Join them using a NewLine character so each starts on a new line
            return string.Join(Environment.NewLine, lines);
        }

        public string GetPalpation(dynamic model)
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

        public string NeckCC(AIIntakeFormModel model)
        {
            var sb = new StringBuilder();
            sb.Append("CERVICAL SPINE: ");

            var complaintItems = new List<string> { "pain" };

            if (model.NeckStiffness != null && model.NeckStiffness.Count > 0)
            {
                // Filter out "Diff..." and "rotating..." so they go to the difficulty sentence
                var symptoms = model.NeckStiffness
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Where(s => !s.Trim().StartsWith("Diff", StringComparison.OrdinalIgnoreCase) &&
                                !s.Trim().Contains("rotating", StringComparison.OrdinalIgnoreCase))
                    .Select(s => s.Trim().ToLower())
                    .ToList();

                complaintItems.AddRange(symptoms);
            }

            // Dynamic output:
            // - Only Pain: "The patient complains of pain in the neck region. "
            // - Pain + Stiffness: "The patient complains of pain and stiffness in the neck region. "
            // - Pain + Stiffness + Spasms: "The patient complains of pain, stiffness, and spasms in the neck region. "
            sb.Append($"The patient complains of {FormatList(complaintItems)} in the neck region. ");

            // 2. Radiation & Associated Symptoms
            var radiationParts = new List<string>();
            string sidePrefix = !string.IsNullOrWhiteSpace(model.NeckRadiatesTo)
                ? model.NeckRadiatesTo.Trim().ToLower() + " "
                : "";

            if (model.NeckRadiates != null && model.NeckRadiates.Count > 0)
            {
                for (int i = 0; i < model.NeckRadiates.Count; i++)
                {
                    // Clean punctuation like trailing dots (e.g., "fingers. ")
                    string part = model.NeckRadiates[i].Trim().TrimEnd('.');

                    // Prefix the first radiation location with the side (e.g., "right arm")
                    if (i == 0 && !string.IsNullOrEmpty(sidePrefix))
                    {
                        radiationParts.Add($"{sidePrefix}{part}");
                    }
                    else
                    {
                        radiationParts.Add(part);
                    }
                }
            }

            string radiationText = FormatList(radiationParts);
            string associatedText = FormatList(model.NeckAssociated?.Select(s =>
                s.Trim().Equals("burning", StringComparison.OrdinalIgnoreCase) ? "burning sensation" : s.Trim()).ToList());

            if (!string.IsNullOrEmpty(radiationText) && !string.IsNullOrEmpty(associatedText))
            {
                sb.Append($"The pain radiates from the neck to {radiationText} and associated with {associatedText}. ");
            }
            else if (!string.IsNullOrEmpty(radiationText))
            {
                sb.Append($"The pain radiates from the neck to {radiationText}. ");
            }
            else if (!string.IsNullOrEmpty(associatedText))
            {
                sb.Append($"The pain is associated with {associatedText}. ");
            }

            // 3. Functional Difficulties
            var difficulties = new List<string>();
            if (model.NeckStiffness != null)
            {
                bool diffTurning = model.NeckStiffness.Any(x => x.Contains("Diff turning", StringComparison.OrdinalIgnoreCase));
                bool rotatingHead = model.NeckStiffness.Any(x => x.Contains("rotating head", StringComparison.OrdinalIgnoreCase));
                bool diffGripping = model.NeckStiffness.Any(x => x.Contains("gripping", StringComparison.OrdinalIgnoreCase));

                if (diffTurning && rotatingHead)
                    difficulties.Add("turning/rotating the head");
                else if (diffTurning)
                    difficulties.Add("turning the head");
                else if (rotatingHead)
                    difficulties.Add("rotating the head");

                if (diffGripping)
                    difficulties.Add("gripping objects");
            }

            if (difficulties.Count > 0)
            {
                sb.Append($"The patient has difficulty {FormatList(difficulties)}. ");
            }

            // 4. Physical Therapy Improvement
            if (!string.IsNullOrWhiteSpace(model.NeckSustainedStiffness))
            {
                if (model.NeckSustainedStiffness.Contains("No sustained", StringComparison.OrdinalIgnoreCase))
                    sb.Append("There is no sustained improvement with physical therapy. ");
                else if (model.NeckSustainedStiffness.Contains("Sustained", StringComparison.OrdinalIgnoreCase))
                    sb.Append("There is sustained improvement with physical therapy. ");
            }

            // 5. Aggravating & Relieving Factors (if selected)
            if (model.NeckWorsens != null && model.NeckWorsens.Count > 0)
            {
                sb.Append($"The pain is worsened with {FormatList(model.NeckWorsens.Select(x => x.Trim()).ToList())}. ");
            }

            if (model.NeckImproves != null && model.NeckImproves.Count > 0)
            {
                sb.Append($"The pain is improved with {FormatList(model.NeckImproves.Select(x => x.Trim()).ToList())}. ");
            }

            // 6. Pain Score
            string painScore = !string.IsNullOrWhiteSpace(model.NeckPain) ? model.NeckPain.Trim() : "___";
            sb.Append($"Pain score is {painScore}/10.");

            return sb.ToString().Trim();
        }

        public string MidbackCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature
            // Template: "The patient complains of mid-back pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull, and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.MbPain) ? modal.MbPain.Trim() : modal.MidbackSection;
            var symptomsList = modal.MbSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList.Count == 0)
                symptomsList = modal.mdbackPain?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));

                if (string.IsNullOrEmpty(modal.mdbackSustainedPain))
                    sentences.Add($"The patient complains of mid-back pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
                else
                    sentences.Add($"The patient complains of mid-back pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} {modal.mdbackSustainedPain}.");
            }
            else
            {
                sentences.Add($"The patient complains of mid-back pain that is {painScore}/10, with 10 being the worst.");
            }

            if (!string.IsNullOrEmpty(modal.MidbackRadiatesTo))
                sentences.Add($"Midback pain {modal.MidbackRadiatesTo}.");

            // 2. Worsens With
            // Template: "The mid-back pain worsens with sitting, standing, bending forward, bending backward, sleeping, twisting, and lifting."
            var worsensList = modal.MbWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(NormalizeWorseningTerm)
                .Distinct()
                .ToList();

            if (worsensList.Count == 0)
                worsensList = modal.MidbackWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(NormalizeWorseningTerm)
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"The mid-back pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With
            // Template: "Mid-back pain improves with resting, medication, and therapy."
            var improvesList = modal.MbImproves != null ? new List<string>(modal.MbImproves) : new List<string>();

            if (improvesList.Count == 0)
                improvesList = modal.MidbackImproves != null ? new List<string>(modal.MidbackImproves) : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.MbImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.MbImprovesOther.Trim());
            }
            else if (!string.IsNullOrWhiteSpace(modal.MidbackImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.MidbackImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(NormalizeImprovementTerm)
                    .Distinct()
                    .ToList();

                sentences.Add($"Mid-back pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        private string FormatList(IEnumerable<string> items)
        {
            var list = items?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            if (list == null || !list.Any()) return string.Empty;
            if (list.Count == 1) return list[0];
            if (list.Count == 2) return $"{list[0]} and {list[1]}";
            return $"{string.Join(", ", list.Take(list.Count - 1))}, and {list.Last()}";
        }

        private string PluralizeBodyPart(string part, string laterality = null)
        {
            if (string.IsNullOrWhiteSpace(part)) return part;

            bool isBilateral = string.Equals(laterality?.Trim(), "bilateral", StringComparison.OrdinalIgnoreCase);

            switch (part.Trim().TrimEnd('.').ToLower())
            {
                case "shoulder":
                case "shoulders":
                    return isBilateral ? "shoulders" : "shoulder";

                case "arm":
                case "arms":
                    return isBilateral ? "arms" : "arm";

                case "forearm":
                case "forearms":
                    return isBilateral ? "forearms" : "forearm";

                case "hand":
                case "hands":
                    return isBilateral ? "hands" : "hand";

                case "wrist":
                case "wrists":
                    return isBilateral ? "wrists" : "wrist";

                case "finger":
                case "fingers":
                    return isBilateral ? "fingers" : "finger";

                default:
                    return part;
            }
        }

        private string NormalizeWorseningTerm(string term)
        {
            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "sleep":
                    return "sleeping";
                case "bending backwards":
                    return "bending backward";
                default:
                    return term;
            }
        }

        private string NormalizeImprovementTerm(string term)
        {
            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "sleep":
                    return "sleeping";
                default:
                    return term;
            }
        }

        public string GenerateSocialHistory(AIIntakeFormModel modal)
        {
            var parts = new List<string>();

            // 1. Smoking
            if (string.Equals(modal.Smoke, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(!string.IsNullOrWhiteSpace(modal.PPD)
                    ? $"The patient is a smoker {modal.PPD.Trim()} PPD."
                    : "The patient is a smoker.");
            }
            else
            {
                parts.Add("The patient is a nonsmoker.");
            }

            // 2. Alcohol
            // Since the HTML 'Y' radio button value is "The patient drinks alcohol occasionally."
            if (!string.IsNullOrWhiteSpace(modal.Alcohol) &&
                modal.Alcohol.Contains("drinks alcohol", StringComparison.OrdinalIgnoreCase) &&
                !modal.Alcohol.Contains("does not", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(modal.txtAlcohol))
                {
                    // Inserts txtAlcohol into the blank ("The patient ____ drinks alcohol occasionally.")
                    // Or change to $"The patient drinks alcohol ({txtAlcohol.Trim()})." if txtAlcohol holds quantity/type
                    parts.Add($"The patient {modal.txtAlcohol.Trim()} drinks alcohol occasionally.");
                }
                else
                {
                    parts.Add("The patient drinks alcohol occasionally.");
                }
            }
            else
            {
                parts.Add("The patient does not drink alcohol.");
            }

            // 3. Recreational Drugs
            if (string.Equals(modal.OtherDrugs, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(!string.IsNullOrWhiteSpace(modal.OtherDrugsTxt)
                    ? $"The patient uses recreational drugs {modal.OtherDrugsTxt.Trim()}."
                    : "The patient uses recreational drugs.");
            }
            else
            {
                parts.Add("The patient does not use recreational drugs.");
            }

            return string.Join(" ", parts);
        }

        #region Lowback

        public string LowBackCC(AIIntakeFormModel model)
        {
            var sb = new StringBuilder();
            sb.Append("LUMBAR SPINE: ");

            // 1. Chief Complaint & Stiffness
            bool hasStiffness = model.LbackPain != null &&
                                model.LbackPain.Any(x => x.Contains("Stiffness", StringComparison.OrdinalIgnoreCase));

            if (hasStiffness)
                sb.Append("The patient complains of pain and stiffness in the low back region. ");
            else
                sb.Append("The patient complains of pain in the low back region. ");

            // 2. Radiation & Associated Symptoms
            var radiationParts = new List<string>();
            string sidePrefix = !string.IsNullOrWhiteSpace(model.LowBackRadiatesTo)
                ? model.LowBackRadiatesTo.Trim().ToLower() + " "
                : "";

            if (model.LowBackRadiates != null && model.LowBackRadiates.Count > 0)
            {
                for (int i = 0; i < model.LowBackRadiates.Count; i++)
                {
                    string part = model.LowBackRadiates[i].Trim();
                    // Apply the side (left/right/bilateral) to the first anatomical part (e.g., "right buttock")
                    if (i == 0 && !string.IsNullOrEmpty(sidePrefix))
                    {
                        radiationParts.Add($"{sidePrefix}{part}");
                    }
                    else
                    {
                        radiationParts.Add(part);
                    }
                }
            }

            string radiationText = FormatList(radiationParts);
            string associatedText = FormatList(model.LowBackAssociated?.Select(s =>
                s.Trim().Equals("burning", StringComparison.OrdinalIgnoreCase) ? "burning sensation" : s.Trim()).ToList());

            if (!string.IsNullOrEmpty(radiationText) && !string.IsNullOrEmpty(associatedText))
            {
                sb.Append($"The pain radiates from the low back to {radiationText} and associated with {associatedText}. ");
            }
            else if (!string.IsNullOrEmpty(radiationText))
            {
                sb.Append($"The pain radiates from the low back to {radiationText}. ");
            }
            else if (!string.IsNullOrEmpty(associatedText))
            {
                sb.Append($"The pain is associated with {associatedText}. ");
            }

            // 3. Functional Difficulties
            var difficulties = new List<string>();
            if (model.LbackPain != null)
            {
                if (model.LbackPain.Any(x => x.Contains("sleeping", StringComparison.OrdinalIgnoreCase)))
                    difficulties.Add("sleeping on the back");
                if (model.LbackPain.Any(x => x.Contains("lifting", StringComparison.OrdinalIgnoreCase)))
                    difficulties.Add("lifting objects");
                if (model.LbackPain.Any(x => x.Contains("bending", StringComparison.OrdinalIgnoreCase)))
                    difficulties.Add("bending");
            }

            if (difficulties.Count > 0)
            {
                sb.Append($"The patient has difficulty {FormatList(difficulties)}. ");
            }

            // 4. Physical Therapy Improvement
            if (!string.IsNullOrWhiteSpace(model.LbackSustainedPain))
            {
                if (model.LbackSustainedPain.Contains("No sustained", StringComparison.OrdinalIgnoreCase))
                    sb.Append("There is no sustained improvement with physical therapy. ");
                else if (model.LbackSustainedPain.Contains("Sustained", StringComparison.OrdinalIgnoreCase))
                    sb.Append("There is sustained improvement with physical therapy. ");
            }

            // 5. Worsened / Improved factors (if selected)
            if (model.LowBackWorsens != null && model.LowBackWorsens.Count > 0)
            {
                sb.Append($"The pain is worsened with {FormatList(model.LowBackWorsens.Select(x => x.Trim()).ToList())}. ");
            }

            if (model.LowBackImproves != null && model.LowBackImproves.Count > 0)
            {
                sb.Append($"The pain is improved with {FormatList(model.LowBackImproves.Select(x => x.Trim()).ToList())}. ");
            }

            // 6. Pain Score
            string painScore = !string.IsNullOrWhiteSpace(model.LowBackPain) ? model.LowBackPain.Trim() : "___";
            sb.Append($"Pain score is {painScore}/10.");

            return sb.ToString().Trim();
        }

        /// <summary>
        /// Formats a list of strings into natural English text with Oxford comma:
        /// ["a"] -> "a"
        /// ["a", "b"] -> "a and b"
        /// ["a", "b", "c"] -> "a, b, and c"
        /// </summary>


        /// <summary>
        /// Pluralizes lower body locations (e.g., hip -> hips, foot -> feet).
        /// </summary>
        private string PluralizeLowerBodyPart(string part, string laterality = null)
        {
            if (string.IsNullOrWhiteSpace(part)) return part;

            bool isBilateral = string.Equals(laterality?.Trim(), "bilateral", StringComparison.OrdinalIgnoreCase);

            switch (part.Trim().TrimEnd('.').ToLower())
            {
                case "hip":
                    return isBilateral ? "Hips" : "Hip";

                case "thigh":
                    return isBilateral ? "Thighs" : "Thigh";

                case "knee":
                    return isBilateral ? "Knees" : "Knee";

                case "leg":
                    return isBilateral ? "Legs" : "Leg";

                case "ankle":
                    return isBilateral ? "Ankles" : "Ankle";

                case "foot":
                case "feet":
                    return isBilateral ? "Feet" : "Foot";

                case "toe":
                case "toes":
                    return isBilateral ? "Toes" : "Toe";

                default:
                    return part;
            }
        }


        #endregion

        #region Shoulder
        public string LeftShoulderCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left shoulder pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LShPain) ? modal.LShPain.Trim() : "__";

            var symptomsList = modal.LShSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left shoulder pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left shoulder pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "The shoulder pain worsens with raising the arm, lifting objects, rotation, working and overhead activities."
            var worsensList = modal.LShWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => NormalizeShoulderWorseningTerm(w))
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"The shoulder pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "Shoulder pain improves with resting, medication, and therapy."
            var improvesList = modal.LSHImproves != null
                ? new List<string>(modal.LSHImproves)
                : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.LSHImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.LSHImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => NormalizeShoulderImprovementTerm(i))
                    .Distinct()
                    .ToList();

                sentences.Add($"Shoulder pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        public string RightShoulderCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Output: "The patient complains of right shoulder pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RShPain) ? modal.RShPain.Trim() : "__";

            var symptomsList = modal.RShSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right shoulder pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right shoulder pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Output: "The shoulder pain worsens with raising the arm, lifting objects, rotation, working and overhead activities."
            var worsensList = modal.RShWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => NormalizeShoulderWorseningTerm(w))
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"The shoulder pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Output: "Shoulder pain improves with resting, medication, and therapy."
            var improvesList = modal.RSHImproves != null
                ? new List<string>(modal.RSHImproves)
                : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.RSHImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.RSHImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => NormalizeShoulderImprovementTerm(i))
                    .Distinct()
                    .ToList();

                sentences.Add($"Shoulder pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        #region Helper Functions


        private string NormalizeShoulderWorseningTerm(string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return string.Empty;

            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "raising arm":
                    return "raising the arm";
                case "lifting":
                    return "lifting objects";
                default:
                    return term;
            }
        }

        private string NormalizeShoulderImprovementTerm(string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return string.Empty;

            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "sleep":
                    return "sleeping";
                default:
                    return term;
            }
        }

        #endregion


        #endregion

        #region Knee
        public string RightKneeCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of right knee pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RKnPain) ? modal.RKnPain.Trim() : "__";

            var symptomsList = modal.RKnSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right knee pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right knee pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "The knee pain worsens with squatting, walking, climbing stairs, going downstairs, standing, getting up from chair, and getting out of car."
            var worsensList = modal.RKnWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => NormalizeKneeWorseningTerm(w))
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"The knee pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "The knee pain improves with resting, medication, and therapy."
            var improvesList = modal.RKNImproves != null
                ? new List<string>(modal.RKNImproves)
                : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.RKNImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.RKNImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => NormalizeKneeImprovementTerm(i))
                    .Distinct()
                    .ToList();

                sentences.Add($"The knee pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        public string LeftKneeCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left knee pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LKnPain) ? modal.LKnPain.Trim() : "__";

            var symptomsList = modal.LKnSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left knee pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left knee pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "The knee pain worsens with squatting, walking, climbing stairs, going downstairs, standing, getting up from chair, and getting out of car."
            var worsensList = modal.LKnWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => NormalizeKneeWorseningTerm(w))
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"The knee pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "The knee pain improves with resting, medication, and therapy."
            var improvesList = modal.LKNImproves != null
                ? new List<string>(modal.LKNImproves)
                : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.LKNImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.LKNImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => NormalizeKneeImprovementTerm(i))
                    .Distinct()
                    .ToList();

                sentences.Add($"The knee pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        #region Knee Helper Methods

        private string NormalizeKneeWorseningTerm(string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return string.Empty;

            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "climb stairs":
                    return "climbing stairs";
                case "get up from chair":
                    return "getting up from chair";
                case "get out of car":
                    return "getting out of car";
                default:
                    return term;
            }
        }

        private string NormalizeKneeImprovementTerm(string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return string.Empty;

            term = term.Trim().TrimEnd('.').ToLower();

            switch (term)
            {
                case "sleep":
                    return "sleeping";
                default:
                    return term;
            }
        }

        #endregion
        #endregion

        #region Ankle
        public string RightAnkleCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of right ankle pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RAnkPain) ? modal.RAnkPain.Trim() : "__";

            var symptomsList = modal.RAnkSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right ankle pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right ankle pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Pain specifically at locations
            // Template: "The patient complains of pain specifically at the medial malleolus, lateral malleolus, and achilles tendon."
            var specialLocations = modal.RAnkPainSpecial?
                .Where(loc => !string.IsNullOrWhiteSpace(loc))
                .Select(loc => loc.ToLower().Trim())
                .Distinct()
                .ToList();

            if (specialLocations != null && specialLocations.Any())
            {
                sentences.Add($"The patient complains of pain specifically at the {FormatList(specialLocations)}.");
            }

            return string.Join(" ", sentences);
        }

        public string LeftAnkleCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left ankle pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LAnkPain) ? modal.LAnkPain.Trim() : "__";

            var symptomsList = modal.LAnkSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left ankle pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left ankle pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Pain specifically at locations
            // Template: "The patient complains of pain specifically at the medial malleolus, lateral malleolus, and achilles tendon."
            var specialLocations = modal.LAnkPainSpecial?
                .Where(loc => !string.IsNullOrWhiteSpace(loc))
                .Select(loc => loc.ToLower().Trim())
                .Distinct()
                .ToList();

            if (specialLocations != null && specialLocations.Any())
            {
                sentences.Add($"The patient complains of pain specifically at the {FormatList(specialLocations)}.");
            }

            return string.Join(" ", sentences);
        }
        #endregion

        #region Elbow

        public string RightElbowCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of right elbow pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RElPain) ? modal.RElPain.Trim() : "__";

            var symptomsList = modal.RElSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right elbow pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right elbow pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Elbow pain worsens with movement, raising the arm, lifting objects, rotation and working."
            var worsensList = modal.RElWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Elbow pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "Elbow pain improves with resting, medication, and therapy."
            var improvesList = modal.RElImproves?
                .Where(i => !string.IsNullOrWhiteSpace(i))
                .Select(i => i.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (improvesList != null && improvesList.Any())
            {
                sentences.Add($"Elbow pain improves with {FormatList(improvesList)}.");
            }

            return string.Join(" ", sentences);
        }

        public string LeftElbowCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left elbow pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LElPain) ? modal.LElPain.Trim() : "__";

            var symptomsList = modal.LElSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left elbow pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left elbow pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Elbow pain worsens with movement, raising the arm, lifting objects, rotation and working."
            var worsensList = modal.LElWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Elbow pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "Elbow pain improves with resting, medication, and therapy."
            var improvesList = modal.LElImproves?
                .Where(i => !string.IsNullOrWhiteSpace(i))
                .Select(i => i.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (improvesList != null && improvesList.Any())
            {
                sentences.Add($"Elbow pain improves with {FormatList(improvesList)}.");
            }

            return string.Join(" ", sentences);
        }

        #endregion

        #region Hip
        public string RightHipCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of right hip pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RHipPain) ? modal.RHipPain.Trim() : "__";

            var symptomsList = modal.RHipSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right hip pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right hip pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Hip pain worsens with sitting, standing, movement and activities."
            var worsensList = modal.RHipWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Hip pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "Hip pain improves with resting, medication, and therapy."
            var improvesList = modal.RHipImproves?
                .Where(i => !string.IsNullOrWhiteSpace(i))
                .Select(i => i.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (improvesList != null && improvesList.Any())
            {
                sentences.Add($"Hip pain improves with {FormatList(improvesList)}.");
            }

            return string.Join(" ", sentences);
        }

        public string LeftHipCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left hip pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LHipPain) ? modal.LHipPain.Trim() : "__";

            var symptomsList = modal.LHipSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left hip pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left hip pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Hip pain worsens with sitting, standing, movement and activities."
            var worsensList = modal.LHipWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Hip pain worsens with {FormatList(worsensList)}.");
            }

            // 3. Improves With Sentence
            // Template: "Hip pain improves with resting, medication, and therapy."
            var improvesList = modal.LHipImproves?
                .Where(i => !string.IsNullOrWhiteSpace(i))
                .Select(i => i.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (improvesList != null && improvesList.Any())
            {
                sentences.Add($"Hip pain improves with {FormatList(improvesList)}.");
            }

            return string.Join(" ", sentences);
        }
        #endregion

        #region Wrist
        public string RightWristCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of right wrist pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.RWristPain) ? modal.RWristPain.Trim() : "__";

            var symptomsList = modal.RWristSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of right wrist pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of right wrist pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Wrist pain worsens with lifting objects, rotation, movement and work activities."
            var worsensList = modal.RWristWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Wrist pain worsens with {FormatList(worsensList)}.");
            }

            return string.Join(" ", sentences);
        }

        public string LeftWristCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Template: "The patient complains of left wrist pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LWristPain) ? modal.LWristPain.Trim() : "__";

            var symptomsList = modal.LWristSymptoms?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of left wrist pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of left wrist pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With Sentence
            // Template: "Wrist pain worsens with lifting objects, rotation, movement and work activities."
            var worsensList = modal.LWristWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Wrist pain worsens with {FormatList(worsensList)}.");
            }

            return string.Join(" ", sentences);
        }
        #endregion

        #region Other


        public string GenerateOther(string PainScale, List<string> PainTypes, string title)
        {

            // If no values were entered, return empty
            bool hasScale = !string.IsNullOrWhiteSpace(PainScale);
            bool hasTypes = PainTypes != null && PainTypes.Any();

            if (!hasScale && !hasTypes)
                return string.Empty;

            var sb = new StringBuilder();

            if (hasScale)
            {
                sb.Append($"The patient complains of {title} pain that is {PainScale.Trim()}/10, with 10 being the worst.");
            }

            if (hasTypes)
            {
                var formattedTypes = string.Join(", ", PainTypes.Select(t => t.Trim()));

                if (hasScale)
                    sb.Append($", which is {formattedTypes.ToLower()}.");
                else
                    sb.Append($"Pain is described as {formattedTypes.ToLower()}.");
            }
            else
            {
                sb.Append(".");
            }

            return sb.ToString();
        }


        #endregion

        #region PE

        private string GenerateNeckExamSentence(IEnumerable<string> cervicalPE,
        string palpation,
        string palpationTenderness,
        IEnumerable<string> peTriggerCervical,
        string peMuscleCervical,
        string peMuscleCervicalSide,
        string region = "cervical")
        {
            var lines = new List<string>();

            // 1. Inspection (Checkboxes: CervicalPE)
            var inspectionItems = cervicalPE?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList() ?? new List<string>();

            if (inspectionItems.Count > 0)
            {
                string joined = string.Join(", ", inspectionItems);
                // Capitalize first letter (e.g., "Unremarkable, kyphotic, lordotic.")
                string formatted = char.ToUpper(joined[0]) + joined.Substring(1);
                lines.Add($"Inspection: {formatted}.");
            }

            // 2. Palpation
            var palpationParts = new List<string>();

            // 2a. Tenderness (Radios: Palpation, Palpationtenderness)
            if (!string.IsNullOrWhiteSpace(palpation))
            {
                string severity = palpation.Trim();
                string regionText = string.IsNullOrWhiteSpace(region) ? "" : $" {region.Trim()}";

                // Omit side if tenderness is "No" or no side is selected
                string sideText = (!string.Equals(severity, "No", StringComparison.OrdinalIgnoreCase) &&
                                   !string.IsNullOrWhiteSpace(palpationTenderness))
                    ? $" {palpationTenderness.Trim()}"
                    : "";

                palpationParts.Add(
                    $"{severity} tenderness to palpation of the{regionText} midline and paraspinal muscles{sideText}.");
            }

            // 2b. Trigger points (Checkboxes: PETriggerCervical)
            var triggerItems = peTriggerCervical?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList() ?? new List<string>();

            if (triggerItems.Count > 0)
            {
                palpationParts.Add($"Trigger points noted in {string.Join(", ", triggerItems)}.");
            }

            // 2c. Muscle spasms (Radios: PEMuscleCervical, PEMuscleCervicalSide)
            if (!string.IsNullOrWhiteSpace(peMuscleCervical))
            {
                string spasmSeverity = peMuscleCervical.Trim();

                // Omit side if spasms are "none" or no side is selected
                string spasmSide = (!string.Equals(spasmSeverity, "none", StringComparison.OrdinalIgnoreCase) &&
                                    !string.IsNullOrWhiteSpace(peMuscleCervicalSide))
                    ? $" {peMuscleCervicalSide.Trim()}"
                    : "";

                palpationParts.Add($"Muscle spasms are {spasmSeverity}{spasmSide}.");
            }

            if (palpationParts.Count > 0)
            {
                lines.Add($"Palpation: {string.Join(" ", palpationParts)}");
            }

            return string.Join(Environment.NewLine, lines);
        }

        private string GenerateNerveProvocation(
    string cevicalSpurling,
    IEnumerable<string> cevicalSpurlingTest,
    string cervicalRightTxt,
    string cervicalLeftTxt,
    string cevicalCompression,
    string cevicalCompressionTest)
        {
            var lines = new List<string>();

            // 1. Spurling's test
            if (!string.IsNullOrWhiteSpace(cevicalSpurling))
            {
                string spurlingResult = cevicalSpurling.Trim().ToLower();

                var selectedSides = new HashSet<string>(
                    cevicalSpurlingTest?
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim()) ?? Enumerable.Empty<string>(),
                    StringComparer.OrdinalIgnoreCase);

                bool hasRight = selectedSides.Contains("right");
                bool hasLeft = selectedSides.Contains("left");
                bool hasBilateral = selectedSides.Contains("bilaterally");

                var sideParts = new List<string>();

                if (hasRight)
                {
                    string rightDeg = !string.IsNullOrWhiteSpace(cervicalRightTxt)
                        ? $" at {cervicalRightTxt.Trim()} degrees"
                        : string.Empty;
                    sideParts.Add($"on the right{rightDeg}");
                }

                if (hasLeft)
                {
                    string leftPrefix = hasRight ? "left" : "on the left";
                    string leftDeg = !string.IsNullOrWhiteSpace(cervicalLeftTxt)
                        ? $" at {cervicalLeftTxt.Trim()} degrees"
                        : string.Empty;
                    sideParts.Add($"{leftPrefix}{leftDeg}");
                }

                if (hasBilateral)
                {
                    sideParts.Add("bilaterally");
                }

                string spurlingSidesText = sideParts.Any()
                    ? " " + string.Join(" and ", sideParts)
                    : string.Empty;

                lines.Add($"Spurling test is {spurlingResult}{spurlingSidesText}.");
            }

            // 2. Compression test
            if (!string.IsNullOrWhiteSpace(cevicalCompression))
            {
                string compResult = cevicalCompression.Trim().ToLower();
                string compSideRaw = cevicalCompressionTest?.Trim().ToLower();

                string compSideText = string.Empty;
                if (!string.IsNullOrWhiteSpace(compSideRaw))
                {
                    compSideText = compSideRaw == "bilaterally"
                        ? " bilaterally"
                        : $" on the {compSideRaw}";
                }

                lines.Add($"Compression test is {compResult}{compSideText}.");
            }

            if (!lines.Any())
                return string.Empty;

            return "Nerve Provocation:" + Environment.NewLine + string.Join(Environment.NewLine, lines);
        }

        private string GenerateLowerBackNarrative(IEnumerable<string> lumbarPE,
    string lumbarPalpation,
    string lumbarPalpationMidline,
    IEnumerable<string> peTriggerLumbar,
    string peMuscleLumbar,
    string peMuscleLumbarSide,
    string region = "lumbar")
        {
            var lines = new List<string>();

            // 1. Inspection (Checkboxes: LUMBARPE)
            var inspectionItems = lumbarPE?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList() ?? new List<string>();

            if (inspectionItems.Any())
            {
                string joined = string.Join(", ", inspectionItems);
                string formatted = char.ToUpper(joined[0]) + joined.Substring(1);
                lines.Add($"Inspection: {formatted}.");
            }

            // 2. Palpation
            var palpationParts = new List<string>();

            // 2a. Tenderness (Radios: LUMBARPalpation, LUMBARPalpationmidline)
            if (!string.IsNullOrWhiteSpace(lumbarPalpation))
            {
                string severity = lumbarPalpation.Trim();
                string regionText = string.IsNullOrWhiteSpace(region) ? string.Empty : $" {region.Trim()}";

                // Omit side if tenderness is "No" or no side is selected
                string sideText = (!string.Equals(severity, "No", StringComparison.OrdinalIgnoreCase) &&
                                   !string.IsNullOrWhiteSpace(lumbarPalpationMidline))
                    ? $" {lumbarPalpationMidline.Trim()}"
                    : string.Empty;

                palpationParts.Add(
                    $"{severity} tenderness to palpation of the{regionText} midline and paraspinal muscles{sideText}.");
            }

            // 2b. Trigger points (Checkboxes: PETriggerLUMBAR)
            var triggerItems = peTriggerLumbar?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .ToList() ?? new List<string>();

            if (triggerItems.Any())
            {
                palpationParts.Add($"Trigger points noted in {string.Join(", ", triggerItems)}.");
            }

            // 2c. Muscle spasms (Radios: PEMuscleLUMBAR, PEMuscleLUMBARside)
            if (!string.IsNullOrWhiteSpace(peMuscleLumbar))
            {
                string spasmSeverity = peMuscleLumbar.Trim();

                // Omit side if spasms are "none" or no side is selected
                string spasmSide = (!string.Equals(spasmSeverity, "none", StringComparison.OrdinalIgnoreCase) &&
                                    !string.IsNullOrWhiteSpace(peMuscleLumbarSide))
                    ? $" {peMuscleLumbarSide.Trim()}"
                    : string.Empty;

                palpationParts.Add($"Muscle spasms are {spasmSeverity}{spasmSide}.");
            }

            if (palpationParts.Any())
            {
                lines.Add($"Palpation: {string.Join(" ", palpationParts)}");
            }

            return string.Join(Environment.NewLine, lines);
        }

        private string GenerateLumbarNerveProvocation(
    string lumbarStraight,
    IEnumerable<string> lumbarStraightTest,
    string lumbarStraightRightTxt,
    string lumbarStraightLeftTxt,
    string lumbarFacetloading,
    string lumbarFacetloadingTest)
        {
            var lines = new List<string>();

            // 1. Straight-leg raise
            if (!string.IsNullOrWhiteSpace(lumbarStraight))
            {
                string slrResult = lumbarStraight.Trim().ToLower();

                var selectedSides = new HashSet<string>(
                    lumbarStraightTest?
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim()) ?? Enumerable.Empty<string>(),
                    StringComparer.OrdinalIgnoreCase);

                bool hasRight = selectedSides.Contains("right");
                bool hasLeft = selectedSides.Contains("left");
                bool hasBilateral = selectedSides.Contains("bilaterally");

                var sideParts = new List<string>();

                if (hasRight)
                {
                    string rightDeg = !string.IsNullOrWhiteSpace(lumbarStraightRightTxt)
                        ? $" at {lumbarStraightRightTxt.Trim()} degrees"
                        : string.Empty;
                    sideParts.Add($"on the right{rightDeg}");
                }

                if (hasLeft)
                {
                    string leftPrefix = hasRight ? "left" : "on the left";
                    string leftDeg = !string.IsNullOrWhiteSpace(lumbarStraightLeftTxt)
                        ? $" at {lumbarStraightLeftTxt.Trim()} degrees"
                        : string.Empty;
                    sideParts.Add($"{leftPrefix}{leftDeg}");
                }

                if (hasBilateral)
                {
                    sideParts.Add("bilaterally");
                }

                string slrSidesText = sideParts.Any()
                    ? " " + string.Join(" and ", sideParts)
                    : string.Empty;

                lines.Add($"Straight-leg raise is {slrResult}{slrSidesText}.");
            }

            // 2. Facet loading test
            if (!string.IsNullOrWhiteSpace(lumbarFacetloading))
            {
                string facetResult = lumbarFacetloading.Trim().ToLower();
                string facetSideRaw = lumbarFacetloadingTest?.Trim().ToLower();

                string facetSideText = string.Empty;
                if (!string.IsNullOrWhiteSpace(facetSideRaw))
                {
                    facetSideText = facetSideRaw == "bilaterally"
                        ? " bilaterally"
                        : $" on the {facetSideRaw}";
                }

                lines.Add($"Facet loading test is {facetResult}{facetSideText}.");
            }

            return string.Join(Environment.NewLine, lines);
        }

        #endregion

        #region AI Content

        public string GetCC(AIIntakeFormModel model)
        {
            string pe = "";
            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + this.NeckCC(model);
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.MidbackCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LowBackCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightShoulderCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftShoulderCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightKneeCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftKneeCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightAnkleCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftAnkleCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right Elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightElbowCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftElbowCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightHipCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftHipCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.RightWristCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + this.LeftWristCC(model) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart1))
            {
                pe = pe + "<p>" + this.GenerateOther(model.Other1PainScale, model.Other1PainType, model.OtherBodyPart1) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart2))
            {
                pe = pe + "<p>" + this.GenerateOther(model.Other2PainScale, model.Other2PainType, model.OtherBodyPart2) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart3))
            {
                pe = pe + "<p>" + this.GenerateOther(model.Other3PainScale, model.Other3PainType, model.OtherBodyPart3) + "<br/></p>";
            }
            return pe.Replace(", and", " and");
        }

        public string GetPE(AIIntakeFormModel model, string path)
        {
            string pe = "";

            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                var topSection = this.GenerateNeckExamSentence(model.CervicalPE,
        model.Palpation,
        model.Palpationtenderness,
        model.PETriggerCervical,
        model.PEMuscleCervical,
        model.PEMuscleCervicalSide);

                var nerveSection = this.GenerateNerveProvocation(
   model.CevicalSpurling,
   model.CevicalSpurlingTest,
   model.CERVICALRighttxt,
   model.CERVICALLefttxt,
   model.CevicalCompression,
   model.CevicalCompressionTest
);
                pe = pe + "<br/>" + GetROM(model, "neck", "Neck.html", topSection, path, nerveSection);
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "midback", "Midback.html", "", path);
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                var topSection = this.GenerateLowerBackNarrative(model.LUMBARPE,
                    model.LUMBARPalpation,
                    model.LUMBARPalpationmidline,
                    model.PETriggerLUMBAR,
                    model.PEMuscleLUMBAR,
                    model.PEMuscleLUMBARside);

                var nerveSection = this.GenerateLumbarNerveProvocation(
   model.LumbarStraight,
    model.LumbarStraightTest,
    model.LumbarStraightRighttxt,
    model.LumbarStraightLefttxt,
    model.LumbarFacetloading,
    model.LumbarFacetloadingTest
);

                pe = pe + "<br/>" + GetROM(model, "lowback", "Lowback.html", topSection, path, nerveSection);
            }
            if (model.Complaints.Contains("right shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right shoulder", "Shoulder.html", "", path);
            }
            if (model.Complaints.Contains("left shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left shoulder", "Shoulder.html", "", path);
            }
            if (model.Complaints.Contains("right knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right knee", "Knee.html", "", path);
            }
            if (model.Complaints.Contains("left knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left knee", "Knee.html", "", path);
            }
            if (model.Complaints.Contains("right ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right ankle", "Ankle.html", "", path);
            }
            if (model.Complaints.Contains("left ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left ankle", "Ankle.html", "", path);
            }
            if (model.Complaints.Contains("right Elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right elbow", "Elbow.html", "", path);
            }
            if (model.Complaints.Contains("left elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left elbow", "Elbow.html", "", path);
            }
            if (model.Complaints.Contains("right hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right hip", "Hip.html", "", path);
            }
            if (model.Complaints.Contains("left hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left hip", "Hip.html", "", path);
            }
            if (model.Complaints.Contains("right wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "right wrist", "Wrist.html", "", path);
            }
            if (model.Complaints.Contains("left wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "left wrist", "Wrist.html", "", path);
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart1))
            {
                pe = pe + "<br/><b>" + model.OtherBodyPart1 + "</b><br/>" + GetOtherROM(model.Other1FlexionTitle, model.Other1FlexionROM, model.Other1FlexionNormal,
                    model.Other1ExtTitle, model.Other1ExtRROM, model.Other1ExtRNormal, path);
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart2))
            {
                pe = pe + "<br/><b>" + model.OtherBodyPart2 + "</b><br/>" + GetOtherROM(model.Other2FlexionTitle, model.Other2FlexionROM, model.Other2FlexionNormal,
                    model.Other2ExtTitle, model.Other2ExtRROM, model.Other2ExtRNormal, path);
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart3))
            {
                pe = pe + "<br/><b>" + model.OtherBodyPart3 + "</b><br/>" + GetOtherROM(model.Other3FlexionTitle, model.Other3FlexionROM, model.Other3FlexionNormal,
                    model.Other3ExtTitle, model.Other3ExtRROM, model.Other3ExtRNormal, path);
            }
            return pe;
        }

        public string GetDTR(AIIntakeFormModel model, string path)
        {
            try
            {

                var templateDir = Path.Combine(path, "templates");

                //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
                var filePath = templateDir + "/BHF/DTR.html";


                // 2. Read the entire file content into a string asynchronously
                string htmlString = System.IO.File.ReadAllText(filePath);

                // 2. Parse HTML
                /*var parser = new HtmlParser();
                var document = parser.ParseDocument(htmlString);

                string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
                string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

                string className = "";

                bool isUEPresent = model.Complaints.Any(value =>
                {
                    string sBody = value.Replace("Left", "")
                                        .Replace("Right", "")
                                        .Trim();

                    // In C#, you can do a case-insensitive check directly without calling ToLower()
                    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
                });

                if (!isUEPresent)
                {
                    // 3. Select all elements matching the class and remove them
                    className = "upper"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }
                }

                bool isLEPresent = model.Complaints.Any(value =>
                {
                    string sBody = value.Replace("Left", "")
                                        .Replace("Right", "")
                                        .Trim();

                    // In C#, you can do a case-insensitive check directly without calling ToLower()
                    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
                });

                if (!isLEPresent)
                {
                    // 3. Select all elements matching the class and remove them
                    className = "lower"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }
                }

                if (!isUEPresent && !isLEPresent)
                {  // 3. Select all elements matching the class and remove them
                    className = "hidAll"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }

                }*/

                // 4. Output the result
                //string modifiedHtml = document.Body.InnerHtml; // or document.DocumentElement.OuterHtml
                string modifiedHtml = htmlString; // or document.DocumentElement.OuterHtml

                modifiedHtml = modifiedHtml.Replace("#TricepsLeft", string.IsNullOrEmpty(model.TricepsLeft) ? "2+" : model.TricepsLeft)
                                              .Replace("#TricepsRight", string.IsNullOrEmpty(model.TricepsRight) ? "2+" : model.TricepsRight)
                                              .Replace("#BicepsLeft", string.IsNullOrEmpty(model.BicepsLeft) ? "2+" : model.BicepsLeft)
                                              .Replace("#BicepsRight", string.IsNullOrEmpty(model.BicepsRight) ? "2+" : model.BicepsRight)
                                              .Replace("#BrachioradialisLeft", string.IsNullOrEmpty(model.BrachioradialisLeft) ? "2+" : model.BrachioradialisLeft)
                                              .Replace("#BrachioradialisRight", string.IsNullOrEmpty(model.BrachioradialisRight) ? "2+" : model.BrachioradialisRight)
                                              .Replace("#KneeRELeft", string.IsNullOrEmpty(model.KneeRELeft) ? "2+" : model.KneeRELeft)
                                              .Replace("#KneeRERight", string.IsNullOrEmpty(model.KneeRERight) ? "2+" : model.KneeRERight)
                                              .Replace("#AnkleRELeft", string.IsNullOrEmpty(model.AnkleRELeft) ? "2+" : model.AnkleRELeft)
                                              .Replace("#AnkleRERight", string.IsNullOrEmpty(model.AnkleRERight) ? "2+" : model.AnkleRERight);

                if (!string.IsNullOrEmpty(model.Reflex))
                    modifiedHtml = "<p>" + model.Reflex + "</p>" + modifiedHtml;

                return modifiedHtml;
            }
            catch (Exception ex)
            {
                //this.SaveLog(ex, "");
                return "";
            }

        }

        public string GetSensory(AIIntakeFormModel model, string path)
        {
            string strSE = "";

            if (model.SE != null && model.SE.Any())
            {
                string seItems = model.SE.Count > 1
                    ? $"{string.Join(", ", model.SE.Take(model.SE.Count - 1))} and {model.SE.Last()}"
                    : model.SE.First();

                strSE = $"{seItems}".TrimEnd('.') + ". " + model.Sensory;
            }
            else
            {
                strSE = model.Sensory;
            }


            var templateDir = Path.Combine(path, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/Sensory.html";


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);

            // 2. Parse HTML
            //var parser = new HtmlParser();
            //var document = parser.ParseDocument(htmlString);

            //string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
            //string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

            //string className = "";

            //bool isUEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isUEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "upper"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //bool isLEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isLEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "lower"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //if (!isUEPresent && !isLEPresent)
            //{  // 3. Select all elements matching the class and remove them
            //    className = "hidAll"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }

            //}

            //// 4. Output the result
            //string modifiedHtml = document.Body.InnerHtml;

            htmlString = htmlString.Replace("#UpperShoulderLeft", string.IsNullOrEmpty(model.UpperShoulderLeft) ? "Normal" : model.UpperShoulderLeft)
                                          .Replace("#UpperShoulderRight", string.IsNullOrEmpty(model.UpperShoulderRight) ? "Normal" : model.UpperShoulderRight)
                                          .Replace("#LateralArmLeft", string.IsNullOrEmpty(model.LateralArmLeft) ? "Normal" : model.LateralArmLeft)
                                          .Replace("#LateralArmRight", string.IsNullOrEmpty(model.LateralArmRight) ? "Normal" : model.LateralArmRight)
                                          .Replace("#ForearmC6Left", string.IsNullOrEmpty(model.ForearmC6Left) ? "Normal" : model.ForearmC6Left)
                                          .Replace("#ForearmC6Right", string.IsNullOrEmpty(model.ForearmC6Right) ? "Normal" : model.ForearmC6Right)
                                          .Replace("#MiddleFingerLeft", string.IsNullOrEmpty(model.MiddleFingerLeft) ? "Normal" : model.MiddleFingerLeft)
                                          .Replace("#MiddleFingerRight", string.IsNullOrEmpty(model.MiddleFingerRight) ? "Normal" : model.MiddleFingerRight)
                                          .Replace("#ForearmC8Left", string.IsNullOrEmpty(model.ForearmC8Left) ? "Normal" : model.ForearmC8Left)
                                          .Replace("#ForearmC8Right", string.IsNullOrEmpty(model.ForearmC8Right) ? "Normal" : model.ForearmC8Right)
                                          .Replace("#MedialArmLeft", string.IsNullOrEmpty(model.MedialArmLeft) ? "Normal" : model.MedialArmLeft)
                                          .Replace("#MedialArmRight", string.IsNullOrEmpty(model.MedialArmRight) ? "Normal" : model.MedialArmRight)
                                          .Replace("#CervicalLeft", string.IsNullOrEmpty(model.CervicalLeft) ? "Normal" : model.CervicalLeft)
                                          .Replace("#CervicalRight", string.IsNullOrEmpty(model.CervicalRight) ? "Normal" : model.CervicalRight)
                                          .Replace("#ThighLeft", string.IsNullOrEmpty(model.ThighLeft) ? "Normal" : model.ThighLeft)
                                          .Replace("#ThighRight", string.IsNullOrEmpty(model.ThighRight) ? "Normal" : model.ThighRight)
                                          .Replace("#FootMedialLeft", string.IsNullOrEmpty(model.FootMedialLeft) ? "Normal" : model.FootMedialLeft)
                                          .Replace("#FootMedialRight", string.IsNullOrEmpty(model.FootMedialRight) ? "Normal" : model.FootMedialRight)
                                          .Replace("#FootDorsumLeft", string.IsNullOrEmpty(model.FootDorsumLeft) ? "Normal" : model.FootDorsumLeft)
                                          .Replace("#FootDorsumRight", string.IsNullOrEmpty(model.FootDorsumRight) ? "Normal" : model.FootDorsumRight)
                                          .Replace("#FootLateralLeft", string.IsNullOrEmpty(model.FootLateralLeft) ? "Normal" : model.FootLateralLeft)
                                          .Replace("#FootLateralRight", string.IsNullOrEmpty(model.FootLateralRight) ? "Normal" : model.FootLateralRight)
                                          .Replace("#LumbarLeft", string.IsNullOrEmpty(model.LumbarLeft) ? "Normal" : model.LumbarLeft)
                                          .Replace("#LumbarRight", string.IsNullOrEmpty(model.LumbarRight) ? "Normal" : model.LumbarRight);

            if (!string.IsNullOrEmpty(strSE))
                htmlString = "<p>To " + strSE + "</p><br/>" + htmlString;

            return htmlString;

        }

        public string GetMMST(AIIntakeFormModel model, string path = "")
        {

            var templateDir = Path.Combine(path, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/MMST.html";


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);

            //// 2. Parse HTML
            //var parser = new HtmlParser();
            //var document = parser.ParseDocument(htmlString);

            //string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
            //string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

            //string className = "";

            //bool isUEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isUEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "upper"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //bool isLEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isLEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "lower"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //if (!isUEPresent && !isLEPresent)
            //{  // 3. Select all elements matching the class and remove them
            //    className = "hidAll"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }

            //}

            //// 4. Output the result
            //string modifiedHtml = document.Body.InnerHtml;

            htmlString = htmlString.Replace("#ShoulderAbductionLeft", string.IsNullOrEmpty(model.ShoulderAbductionLeft) ? "5" : model.ShoulderAbductionLeft)
                                          .Replace("#ShoulderAbductionRight", string.IsNullOrEmpty(model.ShoulderAbductionRight) ? "5" : model.ShoulderAbductionRight)
                                          .Replace("#ShoulderFlexionLeft", string.IsNullOrEmpty(model.ShoulderFlexionLeft) ? "5" : model.ShoulderFlexionLeft)
                                          .Replace("#ShoulderFlexionRight", string.IsNullOrEmpty(model.ShoulderFlexionRight) ? "5" : model.ShoulderFlexionRight)
                                          .Replace("#ElbowExtensionLeft", string.IsNullOrEmpty(model.ElbowExtensionLeft) ? "5" : model.ElbowExtensionLeft)
                                          .Replace("#ElbowExtensionRight", string.IsNullOrEmpty(model.ElbowExtensionRight) ? "5" : model.ElbowExtensionRight)
                                          .Replace("#ElbowFlexionLeft", string.IsNullOrEmpty(model.ElbowFlexionLeft) ? "5" : model.ElbowFlexionLeft)
                                          .Replace("#ElbowFlexionRight", string.IsNullOrEmpty(model.ElbowFlexionRight) ? "5" : model.ElbowFlexionRight)
                                          .Replace("#ElbowSupinationLeft", string.IsNullOrEmpty(model.ElbowSupinationLeft) ? "5" : model.ElbowSupinationLeft)
                                          .Replace("#ElbowSupinationRight", string.IsNullOrEmpty(model.ElbowSupinationRight) ? "5" : model.ElbowSupinationRight)
                                          .Replace("#ElbowPronationLeft", string.IsNullOrEmpty(model.ElbowPronationLeft) ? "5" : model.ElbowPronationLeft)
                                          .Replace("#ElbowPronationRight", string.IsNullOrEmpty(model.ElbowPronationRight) ? "5" : model.ElbowPronationRight)
                                          .Replace("#WristFlexionLeft", string.IsNullOrEmpty(model.WristFlexionLeft) ? "5" : model.WristFlexionLeft)
                                          .Replace("#WristFlexionRight", string.IsNullOrEmpty(model.WristFlexionRight) ? "5" : model.WristFlexionRight)
                                          .Replace("#WristExtensionLeft", string.IsNullOrEmpty(model.WristExtensionLeft) ? "5" : model.WristExtensionLeft)
                                          .Replace("#WristExtensionRight", string.IsNullOrEmpty(model.WristExtensionRight) ? "5" : model.WristExtensionRight)
                                          .Replace("#GripLeft", string.IsNullOrEmpty(model.GripLeft) ? "5" : model.GripLeft)
                                          .Replace("#GripRight", string.IsNullOrEmpty(model.GripRight) ? "5" : model.GripRight)
                                          .Replace("#FingerAbductionLeft", string.IsNullOrEmpty(model.FingerAbductionLeft) ? "5" : model.FingerAbductionLeft)
                                          .Replace("#FingerAbductionRight", string.IsNullOrEmpty(model.FingerAbductionRight) ? "5" : model.FingerAbductionRight)
                                          .Replace("#HipFlexionLeft", string.IsNullOrEmpty(model.HipFlexionLeft) ? "5" : model.HipFlexionLeft)
                                          .Replace("#HipFlexionRight", string.IsNullOrEmpty(model.HipFlexionRight) ? "5" : model.HipFlexionRight)
                                          .Replace("#HipAbductionLeft", string.IsNullOrEmpty(model.HipAbductionLeft) ? "5" : model.HipAbductionLeft)
                                          .Replace("#HipAbductionRight", string.IsNullOrEmpty(model.HipAbductionRight) ? "5" : model.HipAbductionRight)
                                          .Replace("#KneeExtensionLeft", string.IsNullOrEmpty(model.KneeExtensionLeft) ? "5" : model.KneeExtensionLeft)
                                          .Replace("#KneeExtensionRight", string.IsNullOrEmpty(model.KneeExtensionRight) ? "5" : model.KneeExtensionRight)
                                          .Replace("#KneeFlexionLeft", string.IsNullOrEmpty(model.KneeFlexionLeft) ? "5" : model.KneeFlexionLeft)
                                          .Replace("#KneeFlexionRight", string.IsNullOrEmpty(model.KneeFlexionRight) ? "5" : model.KneeFlexionRight)
                                          .Replace("#AnkleDorsiLeft", string.IsNullOrEmpty(model.AnkleDorsiLeft) ? "5" : model.AnkleDorsiLeft)
                                          .Replace("#AnkleDorsiRight", string.IsNullOrEmpty(model.AnkleDorsiRight) ? "5" : model.AnkleDorsiRight)
                                          .Replace("#AnklePlantarLeft", string.IsNullOrEmpty(model.AnklePlantarLeft) ? "5" : model.AnklePlantarLeft)
                                          .Replace("#AnklePlantarRight", string.IsNullOrEmpty(model.AnklePlantarRight) ? "5" : model.AnklePlantarRight)
                                          .Replace("#EHLLeft", string.IsNullOrEmpty(model.EHLLeft) ? "5" : model.EHLLeft)
                                          .Replace("#EHLRight", string.IsNullOrEmpty(model.EHLRight) ? "5" : model.EHLRight);

            if (!string.IsNullOrEmpty(model.Motor))
                htmlString = "<p>" + model.Motor + "</p>" + htmlString;


            return htmlString;

        }

        private string GetROM(AIIntakeFormModel model, string bodyPart, string fileName, string topSection = "", string path = "", string nerveSection = "")
        {
            string rom = "";

            var templateDir = Path.Combine(path, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/IMNPFHPC/" + fileName;


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);


            switch (bodyPart.ToLower())
            {
                case "neck":
                    htmlString = htmlString.Replace("#CervicalFlexion", model.CervicalFlexion)
                                           .Replace("#CervicalExtension", model.CervicalExtension)
                                           .Replace("#CervicalLeftRotation", model.CervicalLeftRotation)
                                           .Replace("#CervicalRightRotation", model.CervicalRightRotation)
                                           .Replace("#CervicalLeftLateralbending", model.CervicalLeftLateralbending)
                                           .Replace("#CervicalrightLateralbending", model.CervicalrightLateralbending)
                                           .Replace("#CERVICALRightDeltoid", model.CERVICALRightDeltoid)
                                           .Replace("#CERVICALRightBiceps", model.CERVICALRightBiceps)
                                           .Replace("#CERVICALRightTriceps", model.CERVICALRightTriceps)
                                           .Replace("#CERVICALRightWristext", model.CERVICALRightWristext)
                                           .Replace("#CERVICALRightWristflex", model.CERVICALRightWristflex)
                                           .Replace("#CERVICALRightIntrinsic", model.CERVICALRightIntrinsic)
                                           .Replace("#CERVICALLEFTDeltoid", model.CERVICALLEFTDeltoid)
                                           .Replace("#CERVICALLEFTBiceps", model.CERVICALLEFTBiceps)
                                           .Replace("#CERVICALLEFTTriceps", model.CERVICALLEFTTriceps)
                                           .Replace("#CERVICALLeftWristext", model.CERVICALLeftWristext)
                                           .Replace("#CERVICALLEFTWristflex", model.CERVICALLEFTWristflex)
                                           .Replace("#CERVICALLEFTIntrinsic", model.CERVICALLEFTIntrinsic)
                                           .Replace("#CERVICALRightC5", model.CERVICALRightC5)
                                           .Replace("#CERVICALRightC6", model.CERVICALRightC6)
                                           .Replace("#CERVICALRightC7", model.CERVICALRightC7)
                                           .Replace("#CERVICALRightC8", model.CERVICALRightC8)
                                           .Replace("#CERVICALRightT1", model.CERVICALRightT1)
                                           .Replace("#CERVICALLeftC5", model.CERVICALLeftC5)
                                           .Replace("#CERVICALLeftC6", model.CERVICALLeftC6)
                                           .Replace("#CERVICALLeftC7", model.CERVICALLeftC7)
                                           .Replace("#CERVICALLeftC8", model.CERVICALLeftC8)
                                           .Replace("#CERVICALLeftT1", model.CERVICALLeftT1)
                                           .Replace("#THORACICBicepstendonright", this.ConvertGradeText(model.THORACICBicepstendonright))
                                           .Replace("#THORACICBicepstendonleft", this.ConvertGradeText(model.THORACICBicepstendonleft))
                                           .Replace("#THORACICTricepstendonright", this.ConvertGradeText(model.THORACICTricepstendonright))
                                           .Replace("#THORACICTricepstendonleft", this.ConvertGradeText(model.THORACICTricepstendonleft))
                                           .Replace("#topSection", topSection)
                                           .Replace("#nerveSection", nerveSection);

                    break;
                case "midback":
                    htmlString = htmlString.Replace("#MbFFROM", model.MbFFROM)
                                           .Replace("#MbFFLeft", model.MbFFLeft)
                                           .Replace("#MbFFRight", model.MbFFRight)
                                           .Replace("#MbExt", model.MbExt);
                    break;
                case "lowback":
                    htmlString = htmlString.Replace("#LUMBARFlexion", model.LUMBARFlexion)
                                           .Replace("#LUMBARExtension", model.LUMBARExtension)
                                           .Replace("#LUMBARLeftLateralbending", model.LUMBARLeftLateralbending)
                                           .Replace("#LUMBARightLateralbending", model.LUMBARightLateralbending)
                                           .Replace("#LUMBARRightIliopsoas", model.LUMBARRightIliopsoas)
                                           .Replace("#LUMBARRightQuadriceps", model.LUMBARRightQuadriceps)
                                           .Replace("#LUMBARRightHamstrings", model.LUMBARRightHamstrings)
                                           .Replace("#LUMBARRightTibant", model.LUMBARRightTibant)
                                           .Replace("#LUMBARRightEHL", model.LUMBARRightEHL)
                                           .Replace("#LUMBARRightGS", model.LUMBARRightGS)
                                           .Replace("#LUMBARLEFTIliopsoas", model.LUMBARLEFTIliopsoas)
                                           .Replace("#LUMBARLEFTQuadriceps", model.LUMBARLEFTQuadriceps)
                                           .Replace("#LUMBARLEFTHamstrings", model.LUMBARLEFTHamstrings)
                                           .Replace("#LUMBARLEFTTibant", model.LUMBARLEFTTibant)
                                           .Replace("#LUMBARLEFTEHL", model.LUMBARLEFTEHL)
                                           .Replace("#LUMBARLEFTGS", model.LUMBARLEFTGS)
                                           .Replace("#LUMBARRightL2", model.LUMBARRightL2)
                                           .Replace("#LUMBARRightL3", model.LUMBARRightL3)
                                           .Replace("#LUMBARRightL4", model.LUMBARRightL4)
                                           .Replace("#LUMBARRightL5", model.LUMBARRightL5)
                                           .Replace("#LUMBARRightS1", model.LUMBARRightS1)
                                           .Replace("#LUMBARLeftL2", model.LUMBARLeftL2)
                                           .Replace("#LUMBARLeftL3", model.LUMBARLeftL3)
                                           .Replace("#LUMBARLeftL4", model.LUMBARLeftL4)
                                           .Replace("#LUMBARLeftL5", model.LUMBARLeftL5)
                                           .Replace("#LUMBARLeftS1", model.LUMBARLeftS1)
                                           .Replace("#LUMBARPatellartendonright", this.ConvertGradeText(model.LUMBARPatellartendonright))
                                           .Replace("#LUMBARPatellartendonleft", this.ConvertGradeText(model.LUMBARPatellartendonleft))
                                           .Replace("#LUMBARAchillestendonright", this.ConvertGradeText(model.LUMBARAchillestendonright))
                                           .Replace("#LUMBARAchillestendonleft", this.ConvertGradeText(model.LUMBARAchillestendonleft))
                                           .Replace("#topSection", topSection)
                                           .Replace("#nerveSection", nerveSection);
                    break;
                case "left shoulder":
                    htmlString = htmlString.Replace("#AbductionROM", model.LSHAbductionROM)
                                           .Replace("#ExternalRotationROM", model.LSHERROM)
                                           .Replace("#FlexionROM", model.LSHFlexionROM)
                                           .Replace("#InternalRotationROM", model.LSHIRROM)
                                           .Replace("#AdductionROM", model.LSHAdductionROM)
                                           .Replace("#ExtensionROM", model.LSHExtOM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right shoulder":
                    htmlString = htmlString.Replace("#AbductionROM", model.RSHAbductionROM)
                                           .Replace("#ExternalRotationROM", model.RSHERROM)
                                           .Replace("#FlexionROM", model.RSHFlexionROM)
                                           .Replace("#InternalRotationROM", model.RSHIRROM)
                                           .Replace("#AdductionROM", model.RSHAdductionROM)
                                           .Replace("#ExtensionROM", model.RSHExtOM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left knee":
                    htmlString = htmlString.Replace("#FlexionROM", model.LKNFlexionROM)
                                           .Replace("#ExtensionROM", model.LKNExtROM)
                                           .Replace("#FlexionNormal", model.LKNFlexionNormal)
                                           .Replace("#ExtensionNormal", model.LKNExtNormal)
                                           .Replace("#FlexionTitle", model.LKNTitleROM)
                                           .Replace("#ExtensionTitle", model.LKNExtTitleROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right knee":
                    htmlString = htmlString.Replace("#FlexionROM", model.RKNFlexionROM)
                                          .Replace("#ExtensionROM", model.RKNExtROM)
                                          .Replace("#FlexionNormal", model.RKNFlexionNormal)
                                          .Replace("#ExtensionNormal", model.RKNExtNormal)
                                          .Replace("#FlexionTitle", model.RKNTitleROM)
                                          .Replace("#ExtensionTitle", model.RKNExtTitleROM)
                                          .Replace("#side", "RIGHT");
                    break;
                case "left ankle":
                    htmlString = htmlString.Replace("#PlantarFlexionROM", model.LAnkleFlexionROM)
                                           .Replace("#InversionROM", model.LAnkleInversionROM)
                                           .Replace("#DorsiflexionROM", model.LAnkleDorsiflexionROM)
                                           .Replace("#EversionROM", model.LAnkleEversionROM)
                                           .Replace("#side", "LEFT");

                    break;
                case "right ankle":
                    htmlString = htmlString.Replace("#PlantarFlexionROM", model.RAnkleFlexionROM)
                                           .Replace("#InversionROM", model.RAnkleInversionROM)
                                           .Replace("#DorsiflexionROM", model.RAnkleDorsiflexionROM)
                                           .Replace("#EversionROM", model.RAnkleEversionROM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left elbow":
                    htmlString = htmlString.Replace("#FlexionROM", model.LElbowFlexionROM)
                                           .Replace("#ExtensionROM", model.LElbowExtensionROM)
                                           .Replace("#SupinationROM", model.LElbowSupinationROM)
                                           .Replace("#PronationROM", model.LElbowPronationROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right elbow":
                    htmlString = htmlString.Replace("#FlexionROM", model.RElbowFlexionROM)
                                           .Replace("#ExtensionROM", model.RElbowExtensionROM)
                                           .Replace("#SupinationROM", model.RElbowSupinationROM)
                                           .Replace("#PronationROM", model.RElbowPronationROM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left hip":
                    htmlString = htmlString.Replace("#FlexionROM", model.LHipFlexionROM)
                                           .Replace("#AdductionROM", model.LHipExtensionROM)
                                           .Replace("#ExtensionROM", model.LHipExtensionROM)
                                           .Replace("#InternalRotationROM", model.LHipIRotationROM)
                                           .Replace("#AbductionROM", model.LHipAbductionROM)
                                           .Replace("#ExternalRotationROM", model.LHipERotationROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right hip":
                    htmlString = htmlString.Replace("#FlexionROM", model.RHipFlexionROM)
                                           .Replace("#AdductionROM", model.RHipExtensionROM)
                                           .Replace("#ExtensionROM", model.RHipExtensionROM)
                                           .Replace("#InternalRotationROM", model.RHipIRotationROM)
                                           .Replace("#AbductionROM", model.RHipAbductionROM)
                                           .Replace("#ExternalRotationROM", model.RHipERotationROM)
                                            .Replace("#side", "RIGHT");
                    break;
                case "left wrist":
                    htmlString = htmlString.Replace("#FlexionROM", model.RWristFlexionROM)
                                           .Replace("#UlnarDeviationROM", model.RWristUDROM)
                                           .Replace("#ExtensionROM", model.RWristExtensionROM)
                                           .Replace("#RadialDeviationROM", model.RWristRDROM)
                                           .Replace("#side", "LEFT");

                    break;
                case "right wrist":
                    htmlString = htmlString.Replace("#FlexionROM", model.LWristFlexionROM)
                                           .Replace("#UlnarDeviationROM", model.LWristUDROM)
                                           .Replace("#ExtensionROM", model.LWristExtensionROM)
                                           .Replace("#RadialDeviationROM", model.LWristRDROM)
                                             .Replace("#side", "RIGHT");
                    break;
                default:
                    rom = "No ROM data available for this body part.";
                    break;
            }
            return htmlString;
        }

        public string GetOtherROM(string NeckFFROM, string NeckFFLeft, string NeckFFRight,
            string NeckExt, string NeckExtLeft, string NeckExtRight, string path)
        {
            string rom = "";

            var templateDir = Path.Combine(path, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/Other.html";


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);

            htmlString = htmlString.Replace("#FlexionTitle", NeckFFROM)
                                          .Replace("#FlexionROM", NeckFFLeft)
                                          .Replace("#FlexionNormal", NeckFFRight)
                                          .Replace("#ExtTitle", NeckExt)
                                          .Replace("#ExtRROM", NeckExtLeft)
                                          .Replace("#ExtRNormal", NeckExtRight);

            return htmlString;

        }
        #endregion

        #region History

        public string GenerateHistory(AIIntakeFormModel model)
        {
            var sb = new StringBuilder();

            // 1. Age, Dominant Hand, Gender
            string age = !string.IsNullOrWhiteSpace(model.Age)
                ? model.Age.Trim()
                : CalculateAge(Convert.ToDateTime(model.DOB), Convert.ToDateTime(model.DOE));

            string handedness = model.DominantHand?.Trim() ?? "";
            string gender = model.Gender?.Trim() ?? "";

            // 2. Restraint & Role (Driver / Passenger / Pedestrian / Bicyclist)
            string restraintText = model.SeatBelt?.Equals("Yes", StringComparison.OrdinalIgnoreCase) == true
                ? "restrained"
                : "unrestrained";

            string roleSentencePart = model.PatientAType switch
            {
                "Driver" => $"the {restraintText} driver of a vehicle that was",
                "Front Pass" => $"the {restraintText} front passenger of a vehicle that was",
                "Rear Pass" => $"the {restraintText} rear passenger of a vehicle that was",
                "Pedestrian" => "a pedestrian who was",
                "Bicyclist" => "a bicyclist who was",
                _ => $"the {restraintText} occupant of a vehicle that was"
            };

            // 3. Accident Type & DOA
            string accidentType = model.InjuryType?.Equals("WC", StringComparison.OrdinalIgnoreCase) == true
                ? "work-related"
                : "motor vehicle";

            string doaText = string.IsNullOrEmpty(model.DOA) == false
                ? Convert.ToDateTime(model.DOA).ToString("MM/dd/yyyy")
                : "________";

            sb.Append($"The patient is a {age}-year-old {handedness} {gender}, who was {roleSentencePart} involved in a {accidentType} accident on {doaText}. ");

            // 4. Impact / Collision Side (skip if Pedestrian/Bicyclist with no vehicle impact selected)
            var impactParts = model.PatientAccidentType
                .Select(x => x switch
                {
                    "Frontend" => "front end",
                    "Rearend" => "rear end",
                    "Driverside" => "driver side",
                    "Passside" => "passenger side",
                    _ => x
                })
                .ToList();

            if (!string.IsNullOrWhiteSpace(model.OtherCollision))
                impactParts.Add(model.OtherCollision.Trim());

            if (impactParts.Any())
            {
                sb.Append($"The patient's vehicle was impacted on the {FormatList(impactParts)}. ");
            }

            // 5. Airbags
            if (!string.IsNullOrWhiteSpace(model.Airbagsdeployed))
            {
                string airbags = model.Airbagsdeployed.Equals("Yes", StringComparison.OrdinalIgnoreCase)
                    ? "deployed"
                    : "did not deploy";
                sb.Append($"The airbags {airbags}. ");
            }

            // 6. EMS
            if (!string.IsNullOrWhiteSpace(model.EMS))
            {
                string ems = model.EMS.Equals("Yes", StringComparison.OrdinalIgnoreCase)
                    ? "arrived"
                    : "did not arrive";
                sb.Append($"The EMS {ems} on the scene. ");
            }

            // 7. LOC & Bruises
            bool hasLoc = model.LOC?.Equals("Yes", StringComparison.OrdinalIgnoreCase) == true;
            bool hasBruises = model.Bruises?.Equals("Yes", StringComparison.OrdinalIgnoreCase) == true;

            if (hasLoc && hasBruises)
                sb.Append("The patient reports loss of consciousness and bruises. ");
            else if (hasLoc)
                sb.Append("The patient reports loss of consciousness and denies bruises. ");
            else if (hasBruises)
                sb.Append("The patient denies loss of consciousness and reports bruises. ");
            else
                sb.Append("The patient denies loss of consciousness. ");

            // 8. Body Parts / Chief Complaints
            var bodyParts = model.Complaints
                .Select(c => c switch
                {
                    "Neck" => "neck",
                    "Midback" => "mid back",
                    "Lowback" => "low back",
                    _ => c.ToLower()
                })
                .ToList();

            if (!string.IsNullOrWhiteSpace(model.OtherBodyPart))
                bodyParts.Add(model.OtherBodyPart.Trim());

            string bodyPartsText = bodyParts.Any() ? FormatList(bodyParts) : "________";

            sb.Append($"The patient sustained multiple skeletal injuries including injury to {bodyPartsText}. ");

            // 9. Hospital
            if (model.Hospital?.Equals("Yes", StringComparison.OrdinalIgnoreCase) == true)
            {
                string hospName = !string.IsNullOrWhiteSpace(model.hospitalname)
                    ? model.hospitalname.Trim()
                    : "________";

                // Avoid duplicating "Hospital" if the user already typed it
                string hospitalFull = hospName.EndsWith("Hospital", StringComparison.OrdinalIgnoreCase)
                    ? hospName
                    : $"{hospName} Hospital";

                sb.Append($"The patient was taken on an emergent basis to {hospitalFull} where the patient was treated and released. ");
            }
            else if (model.Hospital?.Equals("No", StringComparison.OrdinalIgnoreCase) == true)
            {
                sb.Append("The patient did not go to any hospital that same day. ");
            }

            // 10. Physical Therapy Duration
            // Uses PhysicalTherapyDuration if provided; otherwise estimates from DOA to DOE/Today
            string ptDuration = !string.IsNullOrWhiteSpace(model.PhysicalTherapyDuration)
                ? model.PhysicalTherapyDuration.Trim()
                : EstimateWeeksOrMonths(string.IsNullOrEmpty(model.DOA) ? null : Convert.ToDateTime(model.DOA), Convert.ToDateTime(model.DOE));

            sb.Append($"The patient has been undergoing physical therapy for the past {ptDuration}. ");

            // 11. Closing Evaluation Scope
            sb.Append($"My evaluation is limited to {bodyPartsText} injury sustained in the accident of {doaText}.");

            return sb.ToString();
        }
        public string GeneratePmh(string isPmh, List<string> pmhList, string pmhOther)
        {
            // 1. If "No"
            if (string.Equals(isPmh, "No", StringComparison.OrdinalIgnoreCase))
            {
                return "Noncontributory.";
            }

            // 2. If "Yes"
            if (string.Equals(isPmh, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                var conditions = new List<string>();

                // Add selected checkboxes
                if (pmhList != null && pmhList.Count > 0)
                {
                    conditions.AddRange(pmhList.Where(x => !string.IsNullOrWhiteSpace(x)));
                }

                // Add "Other" if provided
                if (!string.IsNullOrWhiteSpace(pmhOther))
                {
                    conditions.Add(pmhOther.Trim());
                }

                // If "Yes" was selected but nothing was checked or typed
                if (conditions.Count == 0)
                {
                    return "Past medical history is positive, but no conditions were specified.";
                }

                // Single condition
                if (conditions.Count == 1)
                {
                    return $"{conditions[0]}.";
                }

                // Multiple conditions joined with commas and "and"
                string formattedList = string.Join(", ", conditions.Take(conditions.Count - 1))
                                       + " and "
                                       + conditions.Last();

                return $"{formattedList}.";
            }

            return string.Empty;
        }

        public string GeneratePsh(string isSmh, string pshDetails)
        {
            // 1. If "No"
            if (string.Equals(isSmh, "No", StringComparison.OrdinalIgnoreCase))
            {
                return "Noncontributory.";
            }

            // 2. If "Yes"
            if (string.Equals(isSmh, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                // If the user selected "Yes" but left the text box empty
                if (string.IsNullOrWhiteSpace(pshDetails) ||
                    string.Equals(pshDetails.Trim(), "none", StringComparison.OrdinalIgnoreCase))
                {
                    return "Past surgical history is positive, but no details were specified.";
                }

                // Clean up text and ensure it ends with a single period
                string details = pshDetails.Trim().TrimEnd('.');

                return $"{details}.";
            }

            return string.Empty;
        }

        public string GenerateMeds(string meds, string medsDetails)
        {
            // 1. If "No"
            if (string.Equals(meds, "No", StringComparison.OrdinalIgnoreCase))
            {
                return "None.";
            }

            // 2. If "Yes"
            if (string.Equals(meds, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                // If "Yes" was selected but the text field is empty
                if (string.IsNullOrWhiteSpace(medsDetails) ||
                    string.Equals(medsDetails.Trim(), "none", StringComparison.OrdinalIgnoreCase))
                {
                    return "Patient reports taking medications, but none were specified.";
                }

                // Remove trailing periods to avoid double punctuation ("..")
                string cleanDetails = medsDetails.Trim().TrimEnd('.');

                return $"{cleanDetails}.";
            }

            return string.Empty;
        }

        public string GenerateAllergy(string allergy, string allergyDetails)
        {
            // 1. If "No"
            if (string.Equals(allergy, "No", StringComparison.OrdinalIgnoreCase))
            {
                return "NO KNOWN DRUG ALLERGIES.";
            }

            // 2. If "Yes"
            if (string.Equals(allergy, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                // If "Yes" was selected but the text field is empty
                if (string.IsNullOrWhiteSpace(allergyDetails) ||
                    string.Equals(allergyDetails.Trim(), "none", StringComparison.OrdinalIgnoreCase))
                {
                    return "Patient reports drug allergies, but no details were specified.";
                }

                // Remove trailing period to avoid double punctuation
                string cleanDetails = allergyDetails.Trim().TrimEnd('.');

                return $"{cleanDetails}.";
            }

            return string.Empty;
        }


        public string GenerateGait(string gaitRadio, List<string> gaitCheckboxes, string gaitOther)
        {
            // 1. Primary gait (defaults to 'Normal' if nothing selected)
            string primaryGait = string.IsNullOrWhiteSpace(gaitRadio) ? "Normal" : gaitRadio.Trim();

            // 2. Check if cane checkbox was selected
            bool hasCane = gaitCheckboxes != null &&
                           gaitCheckboxes.Any(x => x.Equals("Ambulates with a cane", StringComparison.OrdinalIgnoreCase));

            // 3. Clean up other text
            string otherDetails = !string.IsNullOrWhiteSpace(gaitOther) ? gaitOther.Trim().TrimEnd('.') : string.Empty;

            // Case A: Normal gait with no assistive devices or extra notes
            if (primaryGait.Equals("Normal", StringComparison.OrdinalIgnoreCase) && !hasCane && string.IsNullOrEmpty(otherDetails))
            {
                return "Normal.";
            }

            // Case B: Normal gait but uses a cane
            if (primaryGait.Equals("Normal", StringComparison.OrdinalIgnoreCase) && hasCane && string.IsNullOrEmpty(otherDetails))
            {
                return "Normal, but patient ambulates with a cane.";
            }

            // Case C: Antalgic / Altered or with notes
            var sentenceParts = new List<string>();

            primaryGait = CapitalizeFirstLetter(primaryGait);

            if (!string.IsNullOrEmpty(otherDetails))
            {
                sentenceParts.Add($"{primaryGait} {otherDetails}");
            }
            else
            {
                sentenceParts.Add($"{primaryGait}");
            }

            if (hasCane)
            {
                sentenceParts.Add("Patient ambulates with a cane");
            }

            return string.Join(". ", sentenceParts) + ".";
        }

        public string GenerateGait(string gaitRdo,
    string gaitOther,
    string gaitOtherTxt)
        {
            var sentences = new List<string>();

            // 1. Gait pattern (Radio: GAITrdo)
            if (!string.IsNullOrWhiteSpace(gaitRdo))
            {
                string gait = gaitRdo.Trim().ToLower();
                string article = "aeiou".IndexOf(gait[0]) >= 0 ? "an" : "a";

                sentences.Add($"The patient has {article} {gait} gait pattern.");
            }

            // 2. Ambulates with (Radio: GAITOther + Textbox: GAITOthertxt)
            string device = gaitOther?.Trim().ToLower();
            string extra = gaitOtherTxt?.Trim();

            if (!string.IsNullOrWhiteSpace(device) && !string.IsNullOrWhiteSpace(extra))
            {
                sentences.Add($"The patient ambulates with a {device} ({extra}).");
            }
            else if (!string.IsNullOrWhiteSpace(device))
            {
                sentences.Add($"The patient ambulates with a {device}.");
            }
            else if (!string.IsNullOrWhiteSpace(extra))
            {
                sentences.Add($"The patient ambulates with {extra}.");
            }

            return string.Join(" ", sentences);
        }
        #endregion

        #region ROS

        private readonly List<string> AllRosOptions = new List<string>
        {
            "headaches",
            "chest pain/shortness of breath",
            "abdominal pain",
            "dizziness",
            "nausea",
            "ringing in ears",
            "bladder incontinence",
            "bowel incontinence",
            "seizure history",
            "sleep issues/difficulty",
            "anxiety/depression"

        };

        public string GenerateRos(List<string> selectedRos, string rosOther = null)
        {
            var admits = new List<string>();

            // 1. Gather all selected options
            if (selectedRos != null)
            {
                admits.AddRange(selectedRos.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()));
            }

            // Add custom other symptom if filled
            if (!string.IsNullOrWhiteSpace(rosOther))
            {
                admits.Add(rosOther.Trim());
            }

            // 2. Gather unselected options (Denies)
            var denies = AllRosOptions
                .Where(opt => !admits.Any(a => string.Equals(a, opt, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            // 3. Build sentences
            var parts = new List<string>();

            if (admits.Any())
            {
                parts.Add($"The patient admits to {FormatList(admits)}.");
            }

            if (denies.Any())
            {
                parts.Add($"The patient denies {FormatList(denies)}.");
            }

            return string.Join(" ", parts);
        }

        public string GenerateRosIrina(AIIntakeFormModel model)
        {
            var items = new (string Label, string SystemKey, string Value)[]
 {
            ("GENERAL", "General", model.GeneralROS),
            ("SKIN", "Skin", model.SkinROS),
            ("HEENT", "HEENT", model.HEENTROS),
            ("NECK", "Neck", model.NeckROS),
            ("CARDIOVASCULAR", "Cardiovascular", model.CardiovascularROS),
            ("RESPIRATORY", "Respiratory", model.RespiratoryROS),
            ("GASTROINTESTINAL", "Gastrointestinal", model.GastrointestinalROS),
            ("URINARY", "Urinary", model.UrinaryROS),
            ("PERIPHERAL VASCULAR", "Peripheral vascular", model.PeripheralvascularROS),
            ("MUSCULOSKELETAL", "Musculoskeletal", model.MusculoskeletalROS),
            ("NEUROLOGICAL", "Neurological", model.NeurologicalROS),
            ("ENDOCRINE", "Endocrine", model.EndocrineROS)
 };

            var sb = new StringBuilder();

            foreach (var item in items)
            {
                if (!string.IsNullOrWhiteSpace(item.Value))
                {
                    // Custom user text entered (e.g. "pain, stiffness")
                    string description = item.Value.Trim();
                    if (!description.EndsWith("."))
                    {
                        description += ".";
                    }
                    sb.AppendLine($"{item.Label}: {description}");
                }
            }
            return sb.ToString().TrimEnd();
        }

        // Helper to format items with commas and "and"
        private string FormatList(List<string> items)
        {
            if (items == null || items.Count == 0) return string.Empty;
            if (items.Count == 1) return items[0];
            if (items.Count == 2) return $"{items[0]} and {items[1]}";

            return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        }

        #endregion

        #region ADL
        public string GenerateADL(AIIntakeFormModel model)
        {
            if (model == null) return string.Empty;

            var items = new List<(string Activity, string Perform, string Comment)>
        {
            ("Personal hygiene, grooming, dressing, toileting", model.ADLPersonalHygienePerform, model.ADLPersonalHygieneDifficulty),
            ("Walking, transferring, ambulating", model.ADLWalkingPerform, model.ADLWalkingDifficulty),
            ("Eating independently", model.ADLEatingPerform, model.ADLEatingDifficulty),
            ("Cook, clean, maintain households", model.ADLHouseholdPerform, model.ADLHouseholdDifficulty),
            ("Shop for food", model.ADLShoppingPerform, model.ADLShoppingDifficulty),
            ("Sleeping", model.ADLSleepingPerform, model.ADLSleepingDifficulty)
        };

            var sb = new StringBuilder();

            foreach (var (activity, perform, comment) in items)
            {
                string line = FormatAdlRow(activity, perform, comment);
                if (!string.IsNullOrEmpty(line))
                {
                    sb.AppendLine("<p>" + line + "</p>");
                }
            }

            return sb.ToString().TrimEnd();
        }

        private string FormatAdlRow(string activity, string perform, string rawComment)
        {
            string comment = rawComment?.Trim()?.TrimEnd('.');

            // 1. When Perform (Yes) is selected
            if (string.Equals(perform, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                return !string.IsNullOrEmpty(comment)
                    ? $"{activity} - able to perform {comment}."
                    : $"{activity} - able to perform.";
            }

            // 2. When Not Perform (No) is selected
            if (string.Equals(perform, "No", StringComparison.OrdinalIgnoreCase))
            {
                return !string.IsNullOrEmpty(comment)
                    ? $"{activity} - cannot perform {comment}."
                    : $"{activity} - cannot perform.";
            }

            // 3. Neither radio is checked, but comments exist
            if (!string.IsNullOrEmpty(comment))
            {
                return $"{activity} - {comment}.";
            }

            return null;
        }
        #endregion

        #region Occupation
        public string GenerateOccupation(string occupation, string jobTitle)
        {
            // 1. If "No"
            if (string.IsNullOrEmpty(occupation))
            {
                return string.Empty;
            }
            else
            {

                if (!string.IsNullOrWhiteSpace(jobTitle))
                {
                    return "Patient worked as " + occupation.TrimEnd('.') + ". " + jobTitle;
                }

                return "Patient worked as " + occupation.TrimEnd('.') + ".";
            }

        }


        public string GenerateJob(AIIntakeFormModel model)
        {
            var activities = new List<string>();

            // Helper local function: only adds if a value exists
            void AddIfPresent(string? value, string label)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    activities.Add($"{label} {value.Trim()} hr.");
                }
            }

            AddIfPresent(model.JobSitting, "Sitting");
            AddIfPresent(model.JobStanding, "Standing");
            AddIfPresent(model.JobDriving, "Driving");
            AddIfPresent(model.JobBending, "Bending");
            AddIfPresent(model.WalkHr, "Walking");
            AddIfPresent(model.JobKneeling, "Kneeling");
            AddIfPresent(model.KneelHr, "Stooping");
            AddIfPresent(model.JobCrouching, "Crouching");
            AddIfPresent(model.JobLifting, "Lifting");

            // Lifting weights (only if any are checked)
            if (model.LiftingWeight != null)
            {
                var validWeights = model.LiftingWeight
                    .Where(w => !string.IsNullOrWhiteSpace(w))
                    .Select(w => w.Trim())
                    .ToList();

                if (validWeights.Any())
                {
                    activities.Add($"If lifting, how many pounds: {string.Join(", ", validWeights)}");
                }
            }

            var lines = new List<string>();

            // Line 1: Only add if at least one activity has a value
            if (activities.Any())
            {
                lines.Add($"On your job how many hours a day do you spend: {string.Join(" ", activities)}");
            }

            // Line 2: Reach Overhead / Reach Below Waist
            var reachParts = new List<string>();

            if (!string.IsNullOrWhiteSpace(model.ReachOverhead))
            {
                reachParts.Add($"Reach Overhead {model.ReachOverhead.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(model.ReachBelowWaist))
            {
                reachParts.Add($"Reach Below Waist {model.ReachBelowWaist.Trim()}");
            }

            // Only add if at least one reach value is present
            if (reachParts.Any())
            {
                lines.Add(string.Join("                      ", reachParts));
            }

            return string.Join(Environment.NewLine, lines);
        }

        private string FormatListWithAnd(List<string> items)
        {
            if (items == null || items.Count == 0) return string.Empty;
            if (items.Count == 1) return items[0];
            if (items.Count == 2) return $"{items[0]} and {items[1]}";

            return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        }
        #endregion

        #region Recommandation
        public string GenerateRecommandation(AIIntakeFormModel model)
        {
            if (model == null) return string.Empty;

            var recList = model.RECOMMENDATIONS ?? new List<string>();
            var sb = new StringBuilder();

            // 1. Initial Recommendations (PT, Chiro, Acupuncture, HEP)
            var therapyItems = new[] { "PT", "Chiro", "Acupuncture", "HEP" };
            var selectedTherapy = therapyItems.Where(t => recList.Contains(t)).ToList();
            if (selectedTherapy.Any())
            {
                sb.AppendLine($"Recommend {string.Join("/", selectedTherapy)}.");
            }

            // 2. Recommend to continue (PT, Chiro, Acupuncture, HEP)
            var continueItems = model.RECOMMENDATIONSCONTINUE ?? new List<string>();
            var selectedContinue = therapyItems.Where(t => continueItems.Contains(t)).ToList();
            if (selectedContinue.Any())
            {
                sb.AppendLine($"Recommend to continue {string.Join("/", selectedContinue)}.");
            }

            // 3. Recommend to continue current medications
            if (recList.Contains("Continue current medications"))
            {
                sb.AppendLine("Recommend to continue current medications.");
            }

            // 4. Prescribed Medications with dosages
            var medsSelected = model.currentmedications ?? new List<string>();
            var medDetails = new List<string>();

            if (medsSelected.Contains("Ibuprofen"))
                medDetails.Add($"Ibuprofen {model.Ibuprofen?.Trim()} mg");

            if (medsSelected.Contains("Flexeril"))
                medDetails.Add($"Flexeril {model.Flexeril?.Trim()} mg");

            if (medsSelected.Contains("Lidoderm"))
                medDetails.Add($"Lidoderm {model.Lidoderm?.Trim()}%");

            if (medsSelected.Contains("Lidocaine"))
                medDetails.Add($"Lidocaine {model.Lidocaine?.Trim()}%");

            if (medsSelected.Contains("Cyclobenzaprine"))
                medDetails.Add($"Cyclobenzaprine {model.Cyclobenzaprine?.Trim()} mg");

            if (medDetails.Any() || !string.IsNullOrWhiteSpace(model.PrescribedMedicationText))
            {
                string medsText = FormatListWithAnd(medDetails);
                if (!string.IsNullOrWhiteSpace(model.PrescribedMedicationText))
                {
                    medsText = string.IsNullOrEmpty(medsText)
                        ? model.PrescribedMedicationText.Trim()
                        : $"{medsText}, {model.PrescribedMedicationText.Trim()}";
                }
                sb.AppendLine($"{medsText} prescribed.");
            }

            // 5. Orthopedist referrals
            if (recList.Contains("Refer to orthopedist"))
            {
                string note = string.IsNullOrWhiteSpace(model.Refertoorthopedist)
                    ? "Refer to orthopedist."
                    : $"Refer to orthopedist: {model.Refertoorthopedist.Trim()}.";
                sb.AppendLine(note);
            }

            if (recList.Contains("Follow-up with orthopedist"))
            {
                string note = string.IsNullOrWhiteSpace(model.followuporthopedist)
                    ? "Follow-up with orthopedist."
                    : $"Follow-up with orthopedist: {model.followuporthopedist.Trim()}.";
                sb.AppendLine(note);
            }

            // 6. Schedule EMG
            var emgOptions = new[] { "UE", "LE" };
            var scheduledEmg = emgOptions.Where(e => recList.Contains(e)).ToList();
            if (scheduledEmg.Any() || !string.IsNullOrWhiteSpace(model.ScheduleEMGOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.ScheduleEMGOther) ? $" ({model.ScheduleEMGOther.Trim()})" : "";
                sb.AppendLine($"Schedule EMG: {string.Join("/", scheduledEmg)}{other}.");
            }

            // 7. Executed EMG
            var executedEmgOptions = model.RECOMMENDATIONSExecuted ?? new List<string>();
            var executedEmg = emgOptions.Where(e => executedEmgOptions.Contains(e)).ToList();
            if (executedEmg.Any())
            {
                sb.AppendLine($"Executed EMG: {string.Join("/", executedEmg)}.");
            }

            // 8. Schedule Cervical Injections
            var cInjections = new[] { "CESI", "CMBB", "CFB" }.Where(i => recList.Contains(i)).ToList();
            if (cInjections.Any() || !string.IsNullOrWhiteSpace(model.CervicalInjectionOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.CervicalInjectionOther) ? $" ({model.CervicalInjectionOther.Trim()})" : "";
                sb.AppendLine($"Schedule Cervical Injection: {string.Join(", ", cInjections)}{other}.");
            }

            // 9. Schedule Lumbar Injections
            var lInjections = new[] { "LESI", "LMBB", "LFB", "LPDA", "LED" }.Where(i => recList.Contains(i)).ToList();
            if (lInjections.Any() || !string.IsNullOrWhiteSpace(model.LumbarInjectionOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.LumbarInjectionOther) ? $" ({model.LumbarInjectionOther.Trim()})" : "";
                sb.AppendLine($"Schedule Lumbar Injection: {string.Join(", ", lInjections)}{other}.");
            }

            // 10. Imaging (MRI, CT, XR)
            var mris = new[] { "CMRI", "TMRI", "LMRI", "LSMRI", "RSMRI", "LKMRI", "RKMRI" }.Where(m => recList.Contains(m)).ToList();
            if (mris.Any() || !string.IsNullOrWhiteSpace(model.MRIRequestOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.MRIRequestOther) ? $" ({model.MRIRequestOther.Trim()})" : "";
                sb.AppendLine($"Request MRI: {string.Join(", ", mris)}{other}.");
            }

            var cts = new[] { "CSCT", "TSCT", "LSCT" }.Where(c => recList.Contains(c)).ToList();
            if (cts.Any() || !string.IsNullOrWhiteSpace(model.CTRequestOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.CTRequestOther) ? $" ({model.CTRequestOther.Trim()})" : "";
                sb.AppendLine($"Request CT: {string.Join(", ", cts)}{other}.");
            }

            var xrs = new[] { "CSXR", "TSXR", "LSXR" }.Where(x => recList.Contains(x)).ToList();
            if (xrs.Any() || !string.IsNullOrWhiteSpace(model.XRRequestOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.XRRequestOther) ? $" ({model.XRRequestOther.Trim()})" : "";
                sb.AppendLine($"Request X-Ray: {string.Join(", ", xrs)}{other}.");
            }

            // 11. Executed TPI and Meds
            var tpis = new[] { "CS", "TS", "LS" }.Where(t => recList.Contains(t)).ToList();
            var tpiMeds = new[] { "Lidocaine", "Toradol", "Kenalog", "dexamethasone", "Marcaine", "B12" }
                          .Where(m => recList.Contains(m)).ToList();

            if (tpis.Any() || !string.IsNullOrWhiteSpace(model.TPIOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.TPIOther) ? $" ({model.TPIOther.Trim()})" : "";
                sb.AppendLine($"Executed TPI: {string.Join(", ", tpis)}{other}.");
            }

            if (tpiMeds.Any() || !string.IsNullOrWhiteSpace(model.MedsOther))
            {
                string other = !string.IsNullOrWhiteSpace(model.MedsOther) ? $" ({model.MedsOther.Trim()})" : "";
                sb.AppendLine($"Meds administered: {string.Join(", ", tpiMeds)}{other}.");
            }

            // 12. Additional free-text notes
            if (!string.IsNullOrWhiteSpace(model.RecommandationOther1))
                sb.AppendLine(model.RecommandationOther1.Trim());

            if (!string.IsNullOrWhiteSpace(model.RecommandationOther2))
                sb.AppendLine(model.RecommandationOther2.Trim());

            return sb.ToString().TrimEnd();
        }

        // Helper: Formats list into "Item 1, Item 2 and Item 3"
        //private  string FormatListWithAnd(List<string> items)
        //{
        //    if (items == null || items.Count == 0) return string.Empty;
        //    if (items.Count == 1) return items[0];
        //    if (items.Count == 2) return $"{items[0]} and {items[1]}";
        //    return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        //}
        #endregion

        #region Private Method

        public string CapitalizeFirstLetter(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return char.ToUpperInvariant(input[0]) + input.Substring(1);
        }

        private string CalculateAge(DateTime? dob, DateTime? referenceDate)
        {
            if (!dob.HasValue) return "___";
            var refDate = referenceDate ?? DateTime.Today;
            int age = refDate.Year - dob.Value.Year;
            if (dob.Value.Date > refDate.AddYears(-age)) age--;
            return age.ToString();
        }

        private string EstimateWeeksOrMonths(DateTime? doa, DateTime endDate)
        {
            if (!doa.HasValue || endDate <= doa.Value) return "___ weeks/months";

            int totalDays = (endDate - doa.Value).Days;
            if (totalDays < 30)
            {
                int weeks = Math.Max(1, totalDays / 7);
                return weeks == 1 ? "1 week" : $"{weeks} weeks";
            }

            int months = Math.Max(1, totalDays / 30);
            return months == 1 ? "1 month" : $"{months} months";
        }

        private string ConvertGradeText(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            string normalized = input.Trim().ToLowerInvariant();

            var numberMap = new Dictionary<string, string>
    {
        { "zero", "0" },
        { "one", "1" },
        { "two", "2" },
        { "three", "3" },
        { "four", "4" },
        { "five", "5" }
    };

            var match = Regex.Match(
                normalized,
                @"^(zero|one|two|three|four|five|[0-5])\s*(plus|minus|\+|-)?$"
            );

            if (!match.Success)
                return input.Trim();

            string numberPart = match.Groups[1].Value;
            string modifierPart = match.Groups[2].Value;

            if (numberMap.TryGetValue(numberPart, out string digit))
            {
                numberPart = digit;
            }

            string symbol = string.Empty;
            if (modifierPart == "plus" || modifierPart == "+")
                symbol = "+";
            else if (modifierPart == "minus" || modifierPart == "-")
                symbol = "-";

            return numberPart + symbol;
        }

        #endregion

    }
}
