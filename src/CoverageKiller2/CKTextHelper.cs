using System.Text.RegularExpressions;

namespace CoverageKiller2.DOM
{
    public static class CKTextHelper
    {
        /// <summary>
        /// Returns a “pretty” version of the raw text.
        /// This method:
        /// 1. Replaces cell markers (\a) with tabs.
        /// 2. Converts any standalone carriage returns (\r) into CR+LF (\r\n).
        /// 3. Temporarily protects newline sequences with a placeholder.
        /// 4. Removes extraneous non-printable control characters (except tab).
        /// 5. Restores the newline sequences.
        /// 6. Trims any trailing CR and LF (but preserves trailing tabs).
        /// </summary>
        /// <param name="rawText">The raw text returned by Word.Range.Text.</param>
        /// <returns>The cleaned-up text suitable for human comparison.</returns>
        public static string Pretty(string rawText)
        {

            if (string.IsNullOrEmpty(rawText))
                return rawText;

            // Step 1: Replace cell marker (\a) with a tab.
            string pretty = rawText.Replace("\a", "\t");

            // Step 2: Ensure any standalone \r are replaced with \r\n.
            // This will convert any \r that is not already followed by \n into \r\n.
            pretty = Regex.Replace(pretty, @"\r(?!\n)", "\r\n");

            // Step 3: Use a placeholder to protect CR+LF sequences.
            const string newlinePlaceholder = "~~NEWLINE~~";
            pretty = pretty.Replace("\r\n", newlinePlaceholder);

            // Step 4: Remove any extra non-printable control characters.
            // This regex removes control characters in the ranges \x00-\x08 and \x0B-\x1F.
            // Note: Tab (\x09) is not removed.
            pretty = Regex.Replace(pretty, @"[\x00-\x08\x0B-\x1F]", string.Empty);

            // Step 5: Restore the CR+LF sequences.
            pretty = pretty.Replace(newlinePlaceholder, "\r\n");

            // Step 6: Trim only trailing CR and LF (preserving any trailing tabs).
            return pretty.TrimEnd('\r', '\n');
        }

        public static string Scrunch(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Remove cell marker (\a)
            text = text.Replace("\a", string.Empty);

            // Remove all whitespace (spaces, tabs, newlines, etc.)
            return Regex.Replace(text, @"\s+", string.Empty);
        }

        public static bool ScrunchEquals(string text1, string text2)
        {
            return Scrunch(text1).Equals(Scrunch(text2), System.StringComparison.Ordinal);
        }
    }
}
