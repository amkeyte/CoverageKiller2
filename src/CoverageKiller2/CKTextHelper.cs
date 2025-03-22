using System.Text.RegularExpressions;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Helper class for converting raw Word.Range text into a human-friendly "pretty" text,
    /// and for normalizing text by removing all whitespace.
    /// </summary>
    public static class CKTextHelper
    {
        /// <summary>
        /// Returns a “pretty” version of the raw text.
        /// This method replaces cell markers (\a) with tabs,
        /// preserves Windows-style newlines (CR+LF),
        /// trims trailing control characters,
        /// and removes extraneous non-printable characters.
        /// </summary>
        /// <param name="rawText">The raw text returned by Word.Range.Text.</param>
        /// <returns>The cleaned-up text suitable for human comparison.</returns>
        public static string Pretty(string rawText)
        {
            if (string.IsNullOrEmpty(rawText))
                return rawText;

            // Replace cell marker (\a) with a tab.
            string pretty = rawText.Replace("\a", "\t");

            // Use a placeholder to protect CR+LF sequences.
            const string newlinePlaceholder = "~~NEWLINE~~";
            pretty = pretty.Replace("\r\n", newlinePlaceholder);

            // Remove any extra non-printable control characters.
            // This regex removes control characters in the ranges \x00-\x08 and \x0B-\x1F.
            pretty = Regex.Replace(pretty, @"[\x00-\x08\x0B-\x1F]", string.Empty);

            // Restore the CR+LF sequences.
            pretty = pretty.Replace(newlinePlaceholder, "\r\n");

            return pretty.TrimEnd();
        }

        /// <summary>
        /// Scrunches the given text by removing all whitespace characters.
        /// Consumers can use this normalized version for reliable comparisons.
        /// </summary>
        /// <param name="text">The text to scrunch.</param>
        /// <returns>The scrunched text with all whitespace removed.</returns>
        public static string Scrunch(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Remove all whitespace (spaces, tabs, newlines) from the text.
            return Regex.Replace(text, @"\s+", string.Empty);
        }
    }
}
