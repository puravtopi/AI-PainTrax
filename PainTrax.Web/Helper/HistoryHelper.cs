namespace PainTrax.Web.Helper
{
    public static class HistoryHelper
    {
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

                return $"Past medical history is notable for {formattedList}.";
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

                return $"Current medications include {cleanDetails}.";
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

                return $"Drug allergies include {cleanDetails}.";
            }

            return string.Empty;
        }
    }
}
