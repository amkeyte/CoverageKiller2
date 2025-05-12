using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using Serilog;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Logging

{
    public static class LH
    {
        public static string GetTableTitle(CKTable table, string markerText)
        {
            return GetTableTitle(table.COMTable, markerText);
        }

        /// <summary>
        /// Attempts to find a paragraph immediately preceding the given table
        /// that contains a marker text, used as a title.
        /// </summary>
        /// <param name="table">The Word table to check.</param>
        /// <param name="markerText">The marker text to search for (scrunched).</param>
        /// <returns>The matching paragraph text, or null if not found.</returns>
        public static string GetTableTitle(Word.Table table, string markerText)
        {
            if (table == null || string.IsNullOrWhiteSpace(markerText)) return null;

            Word.Document doc = table.Range.Document;
            int start = table.Range.Start;
            int searchStart = Math.Max(0, start - 100);
            try
            {
                //juts try to grab the few paragraphs right before the table. title tag should be there.
                //Bug 20250426 00026: fixed - set max from 1 to 0; preents table.start = 0 -> out of range

                Word.Range range = doc.Range(searchStart, start);
                Word.Paragraphs paraList = range.Paragraphs;

                string scrunchedTarget = CKTextHelper.Scrunch(markerText);

                for (int i = paraList.Count; i >= 1; i--)
                {
                    Word.Paragraph para = paraList[i];
                    if (para.Range.End >= start) continue;

                    string paraText = para.Range.Text?.Trim();
                    if (string.IsNullOrWhiteSpace(paraText)) continue;

                    string scrunched = CKTextHelper.Scrunch(paraText);
                    if (scrunched.Contains(scrunchedTarget))
                    {
                        return paraText;
                    }
                }

                return null;
            }
            catch
            {
                return "UNKOWN";
            }
        }

        //TODO find a way to use the callerName
        public static Exception LogThrow(Exception exception = null, [CallerMemberName] string callerName = "")
        {
            Log.Error(exception, exception.Message);
            return exception;
        }

        /// <summary>
        /// Formats a sequence of strings into a multiline string with optional line numbering and prefix.
        /// </summary>
        /// <param name="lines">The sequence of strings to format.</param>
        /// <param name="prefix">Optional prefix to apply to each line.</param>
        /// <param name="includeIndex">If true, include line numbers starting at 1.</param>
        /// <returns>A single formatted string.</returns>
        public static string DumpString(this IEnumerable<string> lines, string preamble = "\n", string prefix = "", bool includeIndex = false)
        {
            if (lines == null) return string.Empty;

            var sb = new StringBuilder();
            sb.Append(preamble);
            int i = 1;
            foreach (var line in lines)
            {
                if (includeIndex)
                    sb.AppendLine($"{prefix}{i++:D2}: {line}");
                else
                    sb.AppendLine($"{prefix}{line}");
            }
            return sb.ToString();
        }

        internal static void Debug(string message, [CallerMemberName] string memberCallerName = "")
        {
            Log.Debug($"Caller {memberCallerName} said:  {message}");
        }
    }
}
