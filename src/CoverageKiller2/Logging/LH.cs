using Serilog;
using System;
using System.Runtime.CompilerServices;

namespace CoverageKiller2.Logging

{
    public static class LH
    {

            //juts try to grab the few paragraphs right before the table. title tag should be there.
            //Bug 20250426 00026: fixed - set max from 1 to 0; preents table.start = 0 -> out of range
            var searchRangeStart = Math.Max(0, table.Start - 100);
            var paraList = table.Document.Range(searchRangeStart, table.Start).Paragraphs;

            int start = table.COMRange.Start;

            string scrunchedTarget = CKTextHelper.Scrunch(markerText);

            for (int i = paraList.Count; i >= 1; i--)
            {
                var para = paraList[i];
                if (para.End >= start) continue; // skip paras after or inside the table

                string paraText = para.Text?.Trim();
                if (string.IsNullOrWhiteSpace(paraText)) continue;

                string scrunched = CKTextHelper.Scrunch(paraText);
                if (scrunched.Contains(scrunchedTarget))
                {
                    return paraText;
                }
            }

            return null;
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


        public static Exception LogThrow(Exception exception = null, [CallerMemberName] string callerName = "")
        {
            Log.Error(exception, exception.Message);
            return exception;
        }


        // Log.Debug("TRACE => {class}.{func}() = {pVal1}",
        //nameof(PRMCEFixer),
        //       nameof(FixFloorSectionCriticalPointReportTable),
        //       $"{nameof(fixer)}[Table.{nameof(fixer.Index)} = {fixer.Index}]");

        public enum PP
        {
            Enter,
            Result,
            TestPoint,
            PropertyGet,
            PropertySet,
        }
        public static string TraceCaller(params object[] paramPairs)
        {
            // Validate the length of pairs (must be even)
            if (paramPairs.Length < 4 || paramPairs.Length % 2 != 0)
            {
                throw new ArgumentException("The number of parameters must be even and at least 4 (traceType, message, className, methodName).");
            }

            // Validate traceType
            if (paramPairs[0] == null || paramPairs[0].GetType() != typeof(PP))
            {
                throw new ArgumentException("Bad caller type flag at param1.");
            }

            PP traceType = (PP)paramPairs[0];
            string message = paramPairs[1] as string ?? string.Empty;

            // Shared logic to format trace message
            return FormatTraceMessage(traceType, message, paramPairs);
        }

        private static string FormatTraceMessage(PP traceType, string message, object[] paramPairs)
        {
            string defaultMessage;

            // Standard switch statement instead of switch expression
            switch (traceType)
            {
                case PP.Enter:
                    defaultMessage = "Entering member:";
                    break;
                case PP.Result:
                    defaultMessage = "Member returned:";
                    break;
                case PP.TestPoint:
                    defaultMessage = "Test point:";
                    break;
                case PP.PropertyGet:
                    defaultMessage = "Property returned:";
                    break;
                case PP.PropertySet:
                    defaultMessage = "Property set to:";
                    break;
                default:
                    throw new ArgumentException("Invalid trace type.");
            }

            message = string.IsNullOrEmpty(message) ? defaultMessage : message;

            string className = paramPairs[2].ToString();
            string methodName = paramPairs[3].ToString();
            string formattedPairs = string.Empty;

            // Formatting name-value pairs
            for (int i = 4; i < paramPairs.Length; i += 2)
            {
                string name = paramPairs[i].ToString();
                string value = paramPairs[i + 1].ToString();
                formattedPairs += $"\t\t[{name} = {value}]";
            }
            formattedPairs = string.IsNullOrEmpty(formattedPairs) ? "" : "\n" + formattedPairs;

            return $"TRACE => {className}.{methodName} :: {message}{formattedPairs}";
        }


        public static string ObjectPath(params string[] objectNames)
        {
            if (objectNames == null || objectNames.Length == 0)
                throw new ArgumentException("At least one object name must be provided.", nameof(objectNames));

            return string.Join(".", objectNames);
        }


        private static int pingDepth = 0;
        private const string _UNKNOWN_ = "UNKNOWN";
        public static void Ping(Type caller, [CallerMemberName] string callerName = "")
        {

            Log.Verbose($"{new string('\t', pingDepth++)}-> Ping from {caller?.Name ?? _UNKNOWN_}::{callerName}");
        }

        public static void Ping([CallerMemberName] string callerName = "")
        {

            Log.Verbose($"{new string('\t', pingDepth++)}-> Ping from {_UNKNOWN_}::{callerName}");
        }
        public static void Ping(string message, Type caller, [CallerMemberName] string callerName = "")

        {
            Log.Verbose($"{new string('\t', pingDepth++)}-> Ping from {caller?.Name ?? _UNKNOWN_}::{callerName} --- {message}");
        }
        public static void Pong([CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{new string('\t', --pingDepth)}<- Pong from {_UNKNOWN_}::{callerName}");
        }
        public static void Pong(Type caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{new string('\t', --pingDepth)}<- Pong from {caller?.Name ?? _UNKNOWN_}::{callerName}");
        }

        public static void Pong(string message, Type caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{new string('\t', pingDepth--)}<- Pong from {caller?.Name ?? _UNKNOWN_}::{callerName} --- {message}");
        }

    }
}
