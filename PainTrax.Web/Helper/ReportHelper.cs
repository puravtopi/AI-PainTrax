using PainTrax.Web.ViewModel;
using System.Globalization;
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


        public static string Bhf_NeckCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Format: "The patient complains of neck pain that is __/10, with 10 being the worst, which is [qualities] in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.NeckPain) ? modal.NeckPain.Trim() : "__";
            var symptomsList = modal.NeckSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower()));
                sentences.Add($"The patient complains of neck pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of neck pain that is {painScore}/10, with 10 being the worst.");
            }

            // 1b. Pain Type (Facet / Radicular) & Axial Exacerbation (if selected)
            if (!string.IsNullOrWhiteSpace(modal.PainOption) || (modal.Axial != null && modal.Axial.Any()))
            {
                var axialTriggers = modal.Axial?.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.ToLower().Trim()).ToList();
                string painType = !string.IsNullOrWhiteSpace(modal.PainOption) ? modal.PainOption.Trim() : "Axial pain";

                if (axialTriggers != null && axialTriggers.Any())
                {
                    sentences.Add($"{CapitalizeFirstLetter(painType.ToLowerInvariant())} is exacerbated by {FormatList(axialTriggers)}.");
                }
                else
                {
                    sentences.Add($"Pain is characterized as {painType.ToLower()}.");
                }
            }

            // 2. Radiation Sentence
            // Format: "Neck pain radiates to left/right/bilateral shoulders, arms, forearms, hands, wrists and fingers."
            var radAreas = modal.NeckRadiates != null ? new List<string>(modal.NeckRadiates) : new List<string>();
            if (!string.IsNullOrWhiteSpace(modal.NeckRadiatesOther))
            {
                radAreas.Add(modal.NeckRadiatesOther.Trim());
            }

            if (radAreas.Any())
            {
                string sidePrefix = !string.IsNullOrWhiteSpace(modal.NeckRadiatesTo)
                    ? modal.NeckRadiatesTo.Trim().ToLower() + " "
                    : string.Empty;

                var pluralizedAreas = radAreas.Select(a => PluralizeBodyPart(a.ToLower(), sidePrefix));
                sentences.Add($"Neck pain radiates to {sidePrefix}{FormatList(pluralizedAreas)}.");
            }

            // 3. Associated Symptoms Sentence
            // Format: "Neck pain is associated with numbness and tingling to left/right/bilateral shoulders, arms, forearms, hands, wrists and fingers."
            var assocSymptoms = modal.NeckAssociated != null ? new List<string>(modal.NeckAssociated) : new List<string>();
            if (!string.IsNullOrWhiteSpace(modal.NeckAssociatedOther))
            {
                assocSymptoms.Add(modal.NeckAssociatedOther.Trim());
            }

            if (assocSymptoms.Any())
            {
                // Separate sensory symptom keywords (numbness, tingling) from body areas (arms, hands, etc.)
                var sensations = assocSymptoms
                    .Where(s => s.Equals("numbness", StringComparison.OrdinalIgnoreCase) ||
                                s.Equals("tingling", StringComparison.OrdinalIgnoreCase))
                    .Select(s => s.ToLower())
                    .Distinct()
                    .ToList();


                string sidePrefix = !string.IsNullOrWhiteSpace(modal.NeckAssociatedTo)
                  ? modal.NeckAssociatedTo.Trim().ToLower() + " "
                  : string.Empty;

                var bodyParts = assocSymptoms
                    .Where(s => !s.Equals("numbness", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("tingling", StringComparison.OrdinalIgnoreCase))
                    .Select(s => PluralizeBodyPart(s.ToLower(), sidePrefix))
                    .Distinct()
                    .ToList();



                if (sensations.Any() && bodyParts.Any())
                {
                    sentences.Add($"Neck pain is associated with {FormatList(sensations)} to {sidePrefix}{FormatList(bodyParts)}.");
                }
                else if (sensations.Any())
                {
                    sentences.Add($"Neck pain is associated with {FormatList(sensations)}.");
                }
                else
                {
                    sentences.Add($"Neck pain is associated with symptoms in the {sidePrefix}{FormatList(bodyParts)}.");
                }
            }

            // 4. Worsens With Sentence
            // Format: "Neck pain worsens with looking up, looking down, turning head to right, turning head to left, driving and twisting."
            var worsensList = modal.NeckWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim().TrimEnd('.').ToLower())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Neck pain worsens with {FormatList(worsensList)}.");
            }

            // 5. Improves With Sentence
            // Format: "Neck pain improves with resting, medication, and therapy."
            var improvesList = modal.NeckImproves != null ? new List<string>(modal.NeckImproves) : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.NeckImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.NeckImprovesOther.Trim());
            }
            else
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
            }

            if (improvesList.Any())
            {
                var cleanImproves = improvesList
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => i.Trim().TrimEnd('.').ToLower())
                    .Distinct()
                    .ToList();

                sentences.Add($"Neck pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        public static string Irina_NeckCC(AIIntakeFormModel model)
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

       

        public static string MidbackCC(AIIntakeFormModel modal)
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


        private static string FormatList(IEnumerable<string> items)
        {
            var list = items?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            if (list == null || !list.Any()) return string.Empty;
            if (list.Count == 1) return list[0];
            if (list.Count == 2) return $"{list[0]} and {list[1]}";
            return $"{string.Join(", ", list.Take(list.Count - 1))}, and {list.Last()}";
        }

        private static string PluralizeBodyPart(string part, string laterality = null)
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

        private static string NormalizeWorseningTerm(string term)
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

        private static string NormalizeImprovementTerm(string term)
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

        #region Lowback

        public static string Bhf_LowBackCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature
            // Template: "The patient complains of lower back pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull, and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LBPain) ? modal.LBPain.Trim() : "__";
            var symptomsList = modal.LBSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLower().Trim()));
                sentences.Add($"The patient complains of lower back pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of lower back pain that is {painScore}/10, with 10 being the worst.");
            }

            // 1b. Pain Type (Facet / SI Joint / Radicular) & Axial Exacerbation (if selected)
            if (!string.IsNullOrWhiteSpace(modal.LBPainOption) || (modal.LBAxial != null && modal.LBAxial.Any()))
            {
                var axialTriggers = modal.LBAxial?.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.ToLower().Trim()).ToList();
                string painType = !string.IsNullOrWhiteSpace(modal.LBPainOption) ? modal.LBPainOption.Trim() : "Axial pain";

                if (axialTriggers != null && axialTriggers.Any())
                {
                    sentences.Add($"{CapitalizeFirstLetter(painType.ToLowerInvariant())} is exacerbated by {FormatList(axialTriggers)}.");
                }
                else
                {
                    sentences.Add($"Pain is characterized as {painType.ToLower()}.");
                }
            }

            // 2. Radiation Sentence
            // Template: "Lower back pain radiates to left/right/bilateral hips, thighs, knees, legs, ankles, and feet."
            var radAreas = modal.LBRadiates != null ? new List<string>(modal.LBRadiates) : new List<string>();
            if (!string.IsNullOrWhiteSpace(modal.LBRadiatesOther))
            {
                radAreas.Add(modal.LBRadiatesOther.Trim());
            }

            if (radAreas.Any())
            {
                string sidePrefix = !string.IsNullOrWhiteSpace(modal.LBRadiatesTo)
                    ? modal.LBRadiatesTo.Trim().ToLower() + " "
                    : string.Empty;

                var pluralizedAreas = radAreas.Select(a => PluralizeLowerBodyPart(a.ToLower(), sidePrefix).ToLower());
                sentences.Add($"Lower back pain radiates to {sidePrefix}{FormatList(pluralizedAreas)}.");
            }

            // 3. Associated Symptoms Sentence
            // Template: "Lower back pain is associated with numbness and tingling to left/right/bilateral hips, thighs, knees, legs, ankles, and feet."
            var assocSymptoms = modal.LBAssociated != null ? new List<string>(modal.LBAssociated) : new List<string>();
            if (!string.IsNullOrWhiteSpace(modal.LBAssociatedOther))
            {
                assocSymptoms.Add(modal.LBAssociatedOther.Trim());
            }

            if (assocSymptoms.Any())
            {
                // In HTML, tingling checkbox value was "forearm"; we normalize it to "tingling"
                var sensations = assocSymptoms
                    .Where(s => s.Equals("numbness", StringComparison.OrdinalIgnoreCase) ||
                                s.Equals("tingling", StringComparison.OrdinalIgnoreCase) ||
                                s.Equals("forearm", StringComparison.OrdinalIgnoreCase))
                    .Select(s => s.Equals("forearm", StringComparison.OrdinalIgnoreCase) ? "tingling" : s.ToLower())
                    .Distinct()
                    .ToList();

                string sidePrefix = !string.IsNullOrWhiteSpace(modal.LBAssociatedTo)
                  ? modal.LBAssociatedTo.Trim().ToLower() + " "
                  : string.Empty;

                var bodyParts = assocSymptoms
                    .Where(s => !s.Equals("numbness", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("tingling", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("forearm", StringComparison.OrdinalIgnoreCase))
                    .Select(s => PluralizeLowerBodyPart(s.ToLower(), sidePrefix).ToLower())
                    .Distinct()
                    .ToList();



                if (sensations.Any() && bodyParts.Any())
                {
                    sentences.Add($"Lower back pain is associated with {FormatList(sensations)} to {sidePrefix}{FormatList(bodyParts)}.");
                }
                else if (sensations.Any())
                {
                    sentences.Add($"Lower back pain is associated with {FormatList(sensations)}.");
                }
                else
                {
                    sentences.Add($"Lower back pain is associated with symptoms in the {sidePrefix}{FormatList(bodyParts)}.");
                }
            }

            // 4. Worsens With Sentence
            // Template: "Lower back pain worsens with sitting, standing, bending forward, bending backwards, sleeping, twisting right, twisting left, and lifting."
            var worsensList = modal.LBWorsens?
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(NormalizeWorseningTerm)
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Lower back pain worsens with {FormatList(worsensList)}.");
            }

            // 5. Improves With Sentence
            // Template: "The lower back pain improves with resting, medication, and therapy."
            var improvesList = modal.LBImproves != null ? new List<string>(modal.LBImproves) : new List<string>();

            if (!string.IsNullOrWhiteSpace(modal.LBImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.LBImprovesOther.Trim());
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

                sentences.Add($"The lower back pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        public static string Irina_LowBackCC(AIIntakeFormModel model)
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
        private static string PluralizeLowerBodyPart(string part, string laterality = null)
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
        public static string LeftShoulderCC(AIIntakeFormModel modal)
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

        public static string RightShoulderCC(AIIntakeFormModel modal)
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


        private static string NormalizeShoulderWorseningTerm(string term)
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

        private static string NormalizeShoulderImprovementTerm(string term)
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
        public static string RightKneeCC(AIIntakeFormModel modal)
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

        public static string LeftKneeCC(AIIntakeFormModel modal)
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

        private static string NormalizeKneeWorseningTerm(string term)
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

        private static string NormalizeKneeImprovementTerm(string term)
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
        public static string RightAnkleCC(AIIntakeFormModel modal)
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

        public static string LeftAnkleCC(AIIntakeFormModel modal)
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

        public static string RightElbowCC(AIIntakeFormModel modal)
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

        public static string LeftElbowCC(AIIntakeFormModel modal)
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
        public static string RightHipCC(AIIntakeFormModel modal)
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

        public static string LeftHipCC(AIIntakeFormModel modal)
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
        public static string RightWristCC(AIIntakeFormModel modal)
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

        public static string LeftWristCC(AIIntakeFormModel modal)
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


        public static string GenerateOther(string PainScale, List<string> PainTypes, string title)
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

        public static string GenerateNeckExamSentence(AIIntakeFormModel model)
        {
            if (model == null) return string.Empty;

            var sentences = new List<string>();

            // --- 1. Tenderness & Spasms Sentence ---
            var tendernessParts = new List<string>();

            // Format Levels (e.g., "C3/4, C4/5 levels")
            string levelsText = (model.NeckTendernessOptions != null && model.NeckTendernessOptions.Any())
                ? string.Join(", ", model.NeckTendernessOptions.Select(x => x.Trim())) + " levels"
                : "cervical levels";

            // Format Position ("bilateral" -> "bilaterally", "left" -> "on the left", "right" -> "on the right")
            string positionText = "";
            if (!string.IsNullOrWhiteSpace(model.NeckTendernessPosition))
            {
                string pos = model.NeckTendernessPosition.Trim().ToLower();
                positionText = pos switch
                {
                    "bilateral" => "bilaterally",
                    "left" => "on the left",
                    "right" => "on the right",
                    _ => pos
                };
            }

            string tendernessSentence = "";
            if (!string.IsNullOrEmpty(positionText))
            {
                tendernessSentence += $" Patient has pain and tenderness over the {positionText} at {levelsText}";
                tendernessSentence += ".";
            }

            sentences.Add(tendernessSentence);
            sentences.Add(" Facet loading maneuvers are positive.");
            // --- 2. Spurling's Test Sentence ---
            if (!string.IsNullOrWhiteSpace(model.NeckSpurlingTest))
            {
                string spurlingResult = model.NeckSpurlingTest.Trim();
                string spurlingSentence = $"The Spurling's test is {spurlingResult}";

                if (!string.IsNullOrWhiteSpace(model.NeckSpurlingTestOther))
                {
                    spurlingSentence += $" {model.NeckSpurlingTestOther.Trim()}";
                }
                spurlingSentence += ".";
                sentences.Add(spurlingSentence);
            }


            return string.Join(" ", sentences);
        }

        public static string GenerateLowerBackNarrative(AIIntakeFormModel model)
        {
            if (model == null) return string.Empty;

            var lines = new List<string>();

            // --- 1. Tenderness & Facet Joints ---
            string position = model.LBTendernessPosition?.Trim() ?? string.Empty;
            string levels = (model.LBTendernessOptions != null && model.LBTendernessOptions.Any())
                ? string.Join(", ", model.LBTendernessOptions.Select(x => x.Trim()))
                : string.Empty;

            // Combine position and levels with appropriate spacing
            string targetArea = string.Join(" ", new[] { position, " at ", levels }.Where(s => !string.IsNullOrWhiteSpace(s)));
            string targetPrefix = !string.IsNullOrWhiteSpace(targetArea) ? $"{targetArea} " : string.Empty;

            lines.Add($"Patient has pain and tenderness over the {targetPrefix}facet joints. Facet loading maneuvers are positive.");

            // --- 2. SI Joint & Special Signs ---
            bool hasSiJoint = !string.IsNullOrWhiteSpace(model.LBOnexam);
            bool hasSiSigns = model.LBSign != null && model.LBSign.Any();

            if (hasSiJoint || hasSiSigns)
            {
                var siParts = new List<string>();

                if (hasSiJoint)
                {
                    siParts.Add($"On exam, patient has tenderness over the {model.LBOnexam.Trim()} SI joints.");
                }

                if (hasSiSigns)
                {
                    string signs = string.Join(", ", model.LBSign.Select(s => s.Trim()));
                    siParts.Add($"{signs} sign are positive.");
                }

                lines.Add(string.Join(" ", siParts));
            }

            // --- 3. SLR Test ---
            if (!string.IsNullOrWhiteSpace(model.LBSlrTest))
            {
                string slrLine = $"The Spurling's test is {model.LBSlrTest.Trim()}";

                if (!string.IsNullOrWhiteSpace(model.LBSlrTestOther))
                {
                    slrLine += $" {model.LBSlrTestOther.Trim()}";
                }

                lines.Add(slrLine);
            }

            return string.Join(Environment.NewLine, lines);
        }

        #endregion

        #region AI Content

        public static string GetCC(AIIntakeFormModel model)
        {
            string pe = "";
            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                if (model.client_code == "imnpfhpc")
                    pe = pe + ReportHelper.Irina_NeckCC(model);
                else
                    pe = pe + ReportHelper.Bhf_NeckCC(model);
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.MidbackCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                if (model.client_code == "imnpfhpc")
                    pe = pe + "<p>" + ReportHelper.Irina_LowBackCC(model) + "<br/></p>";
                else
                    pe = pe + "<p>" + ReportHelper.Bhf_LowBackCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightShoulderCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftShoulderCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightKneeCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftKneeCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightAnkleCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftAnkleCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right Elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightElbowCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftElbowCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightHipCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftHipCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("right wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.RightWristCC(model) + "<br/></p>";
            }
            if (model.Complaints.Contains("left wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<p>" + ReportHelper.LeftWristCC(model) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart1))
            {
                pe = pe + "<p>" + ReportHelper.GenerateOther(model.Other1PainScale, model.Other1PainType, model.OtherBodyPart1) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart2))
            {
                pe = pe + "<p>" + ReportHelper.GenerateOther(model.Other2PainScale, model.Other2PainType, model.OtherBodyPart2) + "<br/></p>";
            }
            if (!string.IsNullOrEmpty(model.OtherBodyPart3))
            {
                pe = pe + "<p>" + ReportHelper.GenerateOther(model.Other3PainScale, model.Other3PainType, model.OtherBodyPart3) + "<br/></p>";
            }
            return pe.Replace(", and", " and");
        }

        public static string GetPE(AIIntakeFormModel model, string path)
        {
            string pe = "";

            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                var topSection = ReportHelper.GenerateNeckExamSentence(model);
                pe = pe + "<br/>" + GetROM(model, "neck", "Neck.html", topSection, path);
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + GetROM(model, "midback", "Midback.html", "", path);
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                var topSection = ReportHelper.GenerateLowerBackNarrative(model);
                pe = pe + "<br/>" + GetROM(model, "lowback", "Lowback.html", topSection, path);
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

        public static string GetDTR(AIIntakeFormModel model, string path)
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

        public static string GetSensory(AIIntakeFormModel model, string path)
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

        public static string GetMMST(AIIntakeFormModel model, string path = "")
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

        public static string GetROM(AIIntakeFormModel model, string bodyPart, string fileName, string topSection = "", string path = "")
        {
            string rom = "";

            var templateDir = Path.Combine(path, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/" + fileName;


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);


            switch (bodyPart.ToLower())
            {
                case "neck":
                    htmlString = htmlString.Replace("#ForwardFlexionROM", model.NeckFFROM)
                                           .Replace("#LRotationROM", model.NeckFFLeft)
                                           .Replace("#RRotationROM", model.NeckFFRight)
                                           .Replace("#ExtensionROM", model.NeckExt)
                                           .Replace("#LLateralFlexionROM", model.NeckExtLeft)
                                           .Replace("#RLateralFlexionROM", model.NeckExtRight)
                                           .Replace("#RLateralFlexionROM", model.NeckExtRight)
                                           .Replace("#topSection", topSection);
                    break;
                case "midback":
                    htmlString = htmlString.Replace("#MbFFROM", model.MbFFROM)
                                           .Replace("#MbFFLeft", model.MbFFLeft)
                                           .Replace("#MbFFRight", model.MbFFRight)
                                           .Replace("#MbExt", model.MbExt);
                    break;
                case "lowback":
                    htmlString = htmlString.Replace("#LbFFROM", model.LBFFROM)
                                           .Replace("#LbFFLeft", model.LBFFLeft)
                                           .Replace("#LbFFRight", model.LBFFRight)
                                           .Replace("#LbExt", model.LBExt)
                                           .Replace("#SLRTest", model.LBSlrTest)
                                            .Replace("#topSection", topSection);
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

        public static string GetOtherROM(string NeckFFROM, string NeckFFLeft, string NeckFFRight,
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

        public static string GenerateHistory(AIIntakeFormModel model)
        {
            if (model == null) return string.Empty;

            var sentences = new List<string>();

            // Determine pronouns
            bool isFemale = string.Equals(model.Gender, "Female", StringComparison.OrdinalIgnoreCase);
            string pronounSubj = isFemale ? "she" : "he";
            string pronounSubjCap = isFemale ? "She" : "He";
            string pronounPoss = isFemale ? "her" : "his";
            string pName = model.LN + " " + model.FN;

            // 1. Encounter / Introduction
            string accidentType = model.StatesType != null && model.StatesType.Any()
                ? string.Join(", ", model.StatesType).Trim()
                : "motor vehicle accident";

            string formattedDoa = !string.IsNullOrWhiteSpace(model.lblDOA) ? model.lblDOA.Trim() : "an unspecified date";
            string title = (isFemale ? "Ms." : "Mr.");
            string patientName = !string.IsNullOrWhiteSpace(pName) ? pName.Trim() : "Patient";
            string handedness = !string.IsNullOrWhiteSpace(model.DominantHand) ? model.DominantHand.Trim() : "right-handed";

            sentences.Add($"On {model.DOE}, {title} {patientName}, a {handedness} {model.Age}-year-old {model.Gender.ToLower()} presents for the initial evaluation of the injuries sustained in a {accidentType} which occurred on {formattedDoa}.");

            // 2. Patient Role & Impact Type
            var patientType = model.PatientType?.FirstOrDefault();
            var vehicleHit = model.VehicleHit?.FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(patientType) || !string.IsNullOrWhiteSpace(vehicleHit))
            {
                string roleText;
                if (!string.IsNullOrWhiteSpace(patientType))
                {
                    if (patientType.Equals("Belted Driver", StringComparison.OrdinalIgnoreCase))
                        roleText = "the restrained driver of a vehicle";
                    else if (patientType.Equals("Pedestrian", StringComparison.OrdinalIgnoreCase))
                        roleText = "a pedestrian";
                    else if (patientType.Equals("Bicyclist", StringComparison.OrdinalIgnoreCase))
                        roleText = "a bicyclist";
                    else
                        roleText = $"a {patientType.ToLower()}";
                }
                else
                {
                    roleText = "an occupant of a vehicle";
                }

                string hitText = !string.IsNullOrWhiteSpace(vehicleHit)
                    ? $" which was involved in a {vehicleHit.ToLower()} collision."
                    : ".";

                sentences.Add($"The patient states {pronounSubj} was {roleText}{hitText}");
            }

            // 3. Free-text Mechanism of Accident (NFDescription)
            if (!string.IsNullOrWhiteSpace(model.NFDescription))
            {
                string desc = model.NFDescription.Trim();
                if (!desc.EndsWith(".")) desc += ".";
                sentences.Add($"Patient states {desc}");
            }

            // 4. Direct Impact
            if (string.Equals(model.DirectImpact, "Yes", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(model.DirectImpactBodyPart))
            {
                sentences.Add($"Due to the impact, {pronounSubj} injured {pronounPoss} {model.DirectImpactBodyPart.Trim()}.");
            }

            // 5. Loss of Consciousness
            if (string.Equals(model.LoseConsciousness, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                string locDuration = !string.IsNullOrWhiteSpace(model.LOCForHowLong)
                    ? $" for {model.LOCForHowLong.Trim()}"
                    : string.Empty;
                sentences.Add($"Patient had loss of consciousness{locDuration}.");
            }

            // 6. EMS
            if (string.Equals(model.EMS, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                sentences.Add("The patient states that an EMS team arrived.");
            }

            // 7. Hospital Visit
            if (string.Equals(model.Hospital, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                string hospName = !string.IsNullOrWhiteSpace(model.HospitalName)
                    ? model.HospitalName.Trim()
                    : "the hospital";

                string transportMethod = !string.IsNullOrWhiteSpace(model.TransportOther)
                    ? model.TransportOther.Trim()
                    : (!string.IsNullOrWhiteSpace(model.Transport) ? model.Transport.Trim() : "");

                string transportStr = !string.IsNullOrWhiteSpace(transportMethod)
                    ? (transportMethod.ToLower().Contains("with") || transportMethod.ToLower().Contains("by")
                        ? transportMethod.ToLower()
                        : $"with {transportMethod.ToLower()}")
                    : "";

                string timingStr = string.Equals(model.SameDay, "Yes", StringComparison.OrdinalIgnoreCase)
                    ? "the same day the accident occurred"
                    : (!string.IsNullOrWhiteSpace(model.DaysLater) ? $"{model.DaysLater.Trim()} days after the accident occurred" : "subsequently");

                string hospSentence = $"Patient went to {hospName} {transportStr} {timingStr}.".Replace("  ", " ");
                sentences.Add(hospSentence);

                // Hospital Tests
                if (string.Equals(model.IsHospitalTest, "Yes", StringComparison.OrdinalIgnoreCase))
                {
                    var tests = model.HospitalTest != null && model.HospitalTest.Any()
                        ? string.Join(" and ", model.HospitalTest)
                        : "diagnostic testing";

                    string testOther = !string.IsNullOrWhiteSpace(model.HospitalTestOther)
                        ? $" {model.HospitalTestOther.Trim()}"
                        : "";

                    sentences.Add($"At the hospital, the patient had {tests}{testOther}.");
                }

                // Hospital Medication
                if (string.Equals(model.IsHospitalMedication, "Yes", StringComparison.OrdinalIgnoreCase))
                {
                    string meds = !string.IsNullOrWhiteSpace(model.HospitalMedication)
                        ? model.HospitalMedication.Trim()
                        : "pain medication";

                    sentences.Add($"At the hospital, prescription was given for {meds}.");
                }
            }
            //else if (string.Equals(model.Hospital, "No", StringComparison.OrdinalIgnoreCase))
            //{
            //    sentences.Add("The patient did not go to the hospital following the accident.");
            //}

            // 8. Past Accident
            if (string.Equals(model.IsAccidentPast, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                string pastDate = !string.IsNullOrWhiteSpace(model.AccidentPast) ? model.AccidentPast.Trim() : "a previous date";
                string injuredParts = !string.IsNullOrWhiteSpace(model.AccidentBodyPart) ? model.AccidentBodyPart.Trim() : "various body parts";
                bool recovered = string.Equals(model.IsAccidentRecover, "Yes", StringComparison.OrdinalIgnoreCase);

                string recoveryStr = recovered ? "has completely recovered" : "has not completely recovered";
                sentences.Add($"Patient had a previous accident on {pastDate} and injured {pronounPoss} {injuredParts} and {recoveryStr}.");
            }

            // 9. Physical Therapy & Chiropractic Care
            bool hadPT = string.Equals(model.PT, "Yes", StringComparison.OrdinalIgnoreCase);
            bool hadChiro = string.Equals(model.Chiro, "Yes", StringComparison.OrdinalIgnoreCase);

            string ptFreq = !string.IsNullOrWhiteSpace(model.PTFreq) ? model.PTFreq.Trim() : "";
            string chiroFreq = !string.IsNullOrWhiteSpace(model.ChiroFreq) ? model.ChiroFreq.Trim() : "";

            if (hadPT && hadChiro)
            {
                sentences.Add($"The patient has been receiving physical therapy {ptFreq}x a week and chiropractic care {chiroFreq}x a week.".Replace("  ", " "));
            }
            else if (hadPT)
            {
                sentences.Add($"The patient has been receiving physical therapy {ptFreq}x a week.".Replace("  ", " "));
            }
            else if (hadChiro)
            {
                sentences.Add($"The patient has been receiving chiropractic care {chiroFreq}x a week.".Replace("  ", " "));
            }

            //// 10. Review of Systems (ROS)
            //if (model.ROS != null && model.ROS.Any())
            //{
            //    string rosList = string.Join(", ", model.ROS);
            //    if (!string.IsNullOrWhiteSpace(model.ROSOther))
            //    {
            //        rosList += $", {model.ROSOther.Trim()}";
            //    }
            //    sentences.Add($"Review of systems is positive for {rosList}.");
            //}

            // Clean multiple whitespace characters and join sentences
            string fullParagraph = string.Join(" ", sentences);
            return Regex.Replace(fullParagraph, @"\s{2,}", " ").Trim();
        }
        public static string GeneratePmh(string isPmh, List<string> pmhList, string pmhOther)
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

        public static string GeneratePsh(string isSmh, string pshDetails)
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

        public static string GenerateMeds(string meds, string medsDetails)
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



        public static string GenerateAllergy(string allergy, string allergyDetails)
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

        public static string GenerateGait(string gaitRadio, List<string> gaitCheckboxes, string gaitOther)
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
                sentenceParts.Add($"{primaryGait} ({otherDetails})");
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
        #endregion

        #region ROS

        private static readonly List<string> AllRosOptions = new List<string>
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

        public static string GenerateRos(List<string> selectedRos, string rosOther = null)
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

        public static string GenerateRosIrina(AIIntakeFormModel model)
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
        private static string FormatList(List<string> items)
        {
            if (items == null || items.Count == 0) return string.Empty;
            if (items.Count == 1) return items[0];
            if (items.Count == 2) return $"{items[0]} and {items[1]}";

            return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        }

        #endregion

        #region ADL
        public static string GenerateADL(AIIntakeFormModel model)
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

        private static string FormatAdlRow(string activity, string perform, string rawComment)
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
        public static string GenerateOccupation(string occupation, string jobTitle)
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
            

        public static string GenerateJob(AIIntakeFormModel model)
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

        private static string FormatListWithAnd(List<string> items)
        {
            if (items == null || items.Count == 0) return string.Empty;
            if (items.Count == 1) return items[0];
            if (items.Count == 2) return $"{items[0]} and {items[1]}";

            return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        }
        #endregion

        #region Recommandation
        public static string GenerateRecommandation(AIIntakeFormModel model)
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
        //private static string FormatListWithAnd(List<string> items)
        //{
        //    if (items == null || items.Count == 0) return string.Empty;
        //    if (items.Count == 1) return items[0];
        //    if (items.Count == 2) return $"{items[0]} and {items[1]}";
        //    return $"{string.Join(", ", items.Take(items.Count - 1))} and {items.Last()}";
        //}
        #endregion

        #region Private Method

        public static string CapitalizeFirstLetter(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return char.ToUpperInvariant(input[0]) + input.Substring(1);
        }

        #endregion

    }
}
