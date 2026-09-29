using PainTrax.Web.ViewModel;
using System.Text;

namespace PainTrax.Web.Helper
{
    public static class CCHelper
    {

        public static string NeckCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature Sentence
            // Format: "The patient complains of neck pain that is __/10, with 10 being the worst, which is [qualities] in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.NeckPain) ? modal.NeckPain.Trim() : "__";
            var symptomsList = modal.NeckSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant()));
                sentences.Add($"The patient complains of neck pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of neck pain that is {painScore}/10, with 10 being the worst.");
            }

            // 1b. Pain Type (Facet / Radicular) & Axial Exacerbation (if selected)
            if (!string.IsNullOrWhiteSpace(modal.PainOption) || (modal.Axial != null && modal.Axial.Any()))
            {
                var axialTriggers = modal.Axial?.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.ToLowerInvariant().Trim()).ToList();
                string painType = !string.IsNullOrWhiteSpace(modal.PainOption) ? modal.PainOption.Trim() : "Axial pain";

                if (axialTriggers != null && axialTriggers.Any())
                {
                    sentences.Add($"{painType} is exacerbated by {FormatList(axialTriggers)}.");
                }
                else
                {
                    sentences.Add($"Pain is characterized as {painType.ToLowerInvariant()}.");
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
                    ? modal.NeckRadiatesTo.Trim().ToLowerInvariant() + " "
                    : string.Empty;

                var pluralizedAreas = radAreas.Select(a => PluralizeBodyPart(a.ToLowerInvariant()));
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
                    .Select(s => s.ToLowerInvariant())
                    .Distinct()
                    .ToList();

                var bodyParts = assocSymptoms
                    .Where(s => !s.Equals("numbness", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("tingling", StringComparison.OrdinalIgnoreCase))
                    .Select(s => PluralizeBodyPart(s.ToLowerInvariant()))
                    .Distinct()
                    .ToList();

                string sidePrefix = !string.IsNullOrWhiteSpace(modal.NeckAssociatedTo)
                    ? modal.NeckAssociatedTo.Trim().ToLowerInvariant() + " "
                    : string.Empty;

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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                    .Select(i => i.Trim().TrimEnd('.').ToLowerInvariant())
                    .Distinct()
                    .ToList();

                sentences.Add($"Neck pain improves with {FormatList(cleanImproves)}.");
            }

            return string.Join(" ", sentences);
        }

        public static string MidbackCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature
            // Template: "The patient complains of mid-back pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull, and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.MbPain) ? modal.MbPain.Trim() : "__";
            var symptomsList = modal.MbSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
                sentences.Add($"The patient complains of mid-back pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of mid-back pain that is {painScore}/10, with 10 being the worst.");
            }

            // 2. Worsens With
            // Template: "The mid-back pain worsens with sitting, standing, bending forward, bending backward, sleeping, twisting, and lifting."
            var worsensList = modal.MbWorsens?
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

            if (!string.IsNullOrWhiteSpace(modal.MbImprovesOther))
            {
                improvesList.RemoveAll(i => i.Equals("Others", StringComparison.OrdinalIgnoreCase));
                improvesList.Add(modal.MbImprovesOther.Trim());
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

        private static string PluralizeBodyPart(string part)
        {
            if (string.IsNullOrWhiteSpace(part)) return part;

            switch (part.ToLowerInvariant())
            {
                case "shoulder": return "shoulders";
                case "arm": return "arms";
                case "forearm": return "forearms";
                case "hand": return "hands";
                case "wrist": return "wrists";
                case "finger": return "fingers";
                default: return part; // preserves already-plural or custom entries like "fingers"
            }
        }

        private static string NormalizeWorseningTerm(string term)
        {
            term = term.Trim().TrimEnd('.').ToLowerInvariant();

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
            term = term.Trim().TrimEnd('.').ToLowerInvariant();

            switch (term)
            {
                case "sleep":
                    return "sleeping";
                default:
                    return term;
            }
        }

        #region Lowback

        public static string LowBackCC(AIIntakeFormModel modal)
        {
            var sentences = new List<string>();

            // 1. Pain Score & Nature
            // Template: "The patient complains of lower back pain that is __/10, with 10 being the worst, which is constant, intermittent, sharp, shooting, electric, throbbing, pulsating, dull, and achy in nature."
            string painScore = !string.IsNullOrWhiteSpace(modal.LBPain) ? modal.LBPain.Trim() : "__";
            var symptomsList = modal.LBSymptoms?.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (symptomsList != null && symptomsList.Any())
            {
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
                sentences.Add($"The patient complains of lower back pain that is {painScore}/10, with 10 being the worst, which is {symptomsStr} in nature.");
            }
            else
            {
                sentences.Add($"The patient complains of lower back pain that is {painScore}/10, with 10 being the worst.");
            }

            // 1b. Pain Type (Facet / SI Joint / Radicular) & Axial Exacerbation (if selected)
            if (!string.IsNullOrWhiteSpace(modal.LBPainOption) || (modal.LBAxial != null && modal.LBAxial.Any()))
            {
                var axialTriggers = modal.LBAxial?.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => a.ToLowerInvariant().Trim()).ToList();
                string painType = !string.IsNullOrWhiteSpace(modal.LBPainOption) ? modal.LBPainOption.Trim() : "Axial pain";

                if (axialTriggers != null && axialTriggers.Any())
                {
                    sentences.Add($"{painType} is exacerbated by {FormatList(axialTriggers)}.");
                }
                else
                {
                    sentences.Add($"Pain is characterized as {painType.ToLowerInvariant()}.");
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
                    ? modal.LBRadiatesTo.Trim().ToLowerInvariant() + " "
                    : string.Empty;

                var pluralizedAreas = radAreas.Select(a => PluralizeLowerBodyPart(a.ToLowerInvariant()));
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
                    .Select(s => s.Equals("forearm", StringComparison.OrdinalIgnoreCase) ? "tingling" : s.ToLowerInvariant())
                    .Distinct()
                    .ToList();

                var bodyParts = assocSymptoms
                    .Where(s => !s.Equals("numbness", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("tingling", StringComparison.OrdinalIgnoreCase) &&
                                !s.Equals("forearm", StringComparison.OrdinalIgnoreCase))
                    .Select(s => PluralizeLowerBodyPart(s.ToLowerInvariant()))
                    .Distinct()
                    .ToList();

                string sidePrefix = !string.IsNullOrWhiteSpace(modal.LBAssociatedTo)
                    ? modal.LBAssociatedTo.Trim().ToLowerInvariant() + " "
                    : string.Empty;

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

        /// <summary>
        /// Pluralizes lower body locations (e.g., hip -> hips, foot -> feet).
        /// </summary>
        private static string PluralizeLowerBodyPart(string part)
        {
            if (string.IsNullOrWhiteSpace(part)) return part;

            switch (part.Trim().TrimEnd('.').ToLowerInvariant())
            {
                case "hip": return "hip";
                case "thigh": return "thigh";
                case "knee": return "knee";
                case "leg": return "leg";
                case "ankle": return "ankle";
                case "foot": return "feet";
                case "toe":
                case "toes": return "toes";
                default: return part;
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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

            term = term.Trim().TrimEnd('.').ToLowerInvariant();

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

            term = term.Trim().TrimEnd('.').ToLowerInvariant();

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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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

            term = term.Trim().TrimEnd('.').ToLowerInvariant();

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

            term = term.Trim().TrimEnd('.').ToLowerInvariant();

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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(loc => loc.ToLowerInvariant().Trim())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(loc => loc.ToLowerInvariant().Trim())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                .Select(i => i.Trim().TrimEnd('.').ToLowerInvariant())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                .Select(i => i.Trim().TrimEnd('.').ToLowerInvariant())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                .Select(i => i.Trim().TrimEnd('.').ToLowerInvariant())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                .Select(i => i.Trim().TrimEnd('.').ToLowerInvariant())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
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
                string symptomsStr = FormatList(symptomsList.Select(s => s.ToLowerInvariant().Trim()));
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
                .Select(w => w.Trim().TrimEnd('.').ToLowerInvariant())
                .Distinct()
                .ToList();

            if (worsensList != null && worsensList.Any())
            {
                sentences.Add($"Wrist pain worsens with {FormatList(worsensList)}.");
            }

            return string.Join(" ", sentences);
        }
        #endregion
    }
}
