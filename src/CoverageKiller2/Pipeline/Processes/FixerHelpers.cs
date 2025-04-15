using System.Text.RegularExpressions;
namespace CoverageKiller2.Pipeline.Processes
{
    internal class FixerHelpers
    {
        public static string NormalizeMatchString(string input)
        {
            return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }
        public bool NormalizeMatchStrings(string str1, string str2)
        {
            return NormalizeMatchString(str1) == NormalizeMatchString(str2);
        }


        /// <summary>
        /// Finds the nearest section break in the specified direction.
        /// </summary>
        /// <param name="range">The range to search from.</param>
        /// <param name="forward">True to search forward, false to search backward.</param>
        /// <returns>The range of the found section break, or null if not found.</returns>
        //public static Word.Range FindSectionBreak(Word.Range range, bool forward)
        //{
        //    Word.Range searchRange = range.Document.Range(
        //        forward ? range.End : 0, // Start search after range for forward, or from start of doc for backward
        //        forward ? range.Document.Content.End : range.Start // Search to end for forward, or up to start for backward
        //    );

        //    searchRange.Find.ClearFormatting();
        //    searchRange.Find.Text = "^b"; // Section break special character
        //    searchRange.Find.Forward = forward;
        //    searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

        //    if (searchRange.Find.Execute())
        //    {
        //        return searchRange; // Return the found section break range
        //    }
        //    return null; // Return null if no section break was found
        //}
    }
}
