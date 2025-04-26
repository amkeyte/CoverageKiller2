using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Logging

{





    /// <summary>
    /// Represents an internal debug exception that signals an unexpected state or logic error
    /// within the CoverageKiller DOM system.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0002
    /// </remarks>
    [Serializable]
    public class CKDebugException : Exception
    {
        /// <inheritdoc/>
        public CKDebugException() { }

        /// <inheritdoc/>
        public CKDebugException(string msg) : base(msg) { }

        /// <inheritdoc/>
        public CKDebugException(string msg, Exception innerException)
            : base(msg, innerException) { }

        /// <inheritdoc/>
        protected CKDebugException(SerializationInfo info, StreamingContext context)
                : base(info, context)
        {
        }
    }
    public static class LH
    {
        public static string GetTableTitle(CKTable table, string markerText)
        {
            if (table == null || string.IsNullOrWhiteSpace(markerText)) return null;

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
            int searchStart = Math.Max(1, start - 100);
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

        /// <summary>
        /// Logs an exception using Serilog and optionally rethrows it.
        /// </summary>
        /// <param name="ex">The exception to log.</param>
        /// <param name="context">Optional context msg to include with the log.</param>
        /// <param name="rethrow">If true, rethrows the exception after logging.</param>
        public static void Error(Exception ex, string context = "", bool rethrow = true)
        {
            Log.Error(ex, "Exception occurred{Context}", string.IsNullOrWhiteSpace(context) ? "" : $" during {context}");

            //if (Debugger.IsAttached) Debugger.Break();

            if (rethrow) throw ex;
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
                throw new ArgumentException("The number of parameters must be even and at least 4 (traceType, msg, className, methodName).");
            }

            // Validate traceType
            if (paramPairs[0] == null || paramPairs[0].GetType() != typeof(PP))
            {
                throw new ArgumentException("Bad caller type flag at param1.");
            }

            PP traceType = (PP)paramPairs[0];
            string msg = paramPairs[1] as string ?? string.Empty;

            // Shared logic to format trace msg
            return FormatTracemsg(traceType, msg, paramPairs);
        }

        private static string FormatTracemsg(PP traceType, string msg, object[] paramPairs)
        {
            string defaultmsg;

            // Standard switch statement instead of switch expression
            switch (traceType)
            {
                case PP.Enter:
                    defaultmsg = "Entering member:";
                    break;
                case PP.Result:
                    defaultmsg = "Member returned:";
                    break;
                case PP.TestPoint:
                    defaultmsg = "Test point:";
                    break;
                case PP.PropertyGet:
                    defaultmsg = "Property returned:";
                    break;
                case PP.PropertySet:
                    defaultmsg = "Property set to:";
                    break;
                default:
                    throw new ArgumentException("Invalid trace type.");
            }

            msg = string.IsNullOrEmpty(msg) ? defaultmsg : msg;

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

            return $"TRACE => {className}.{methodName} :: {msg}{formattedPairs}";
        }


        public static string ObjectPath(params string[] objectNames)
        {
            if (objectNames == null || objectNames.Length == 0)
                throw new ArgumentException("At least one object name must be provided.", nameof(objectNames));

            return string.Join(".", objectNames);
        }
        public static void Checkpoint(string msg, Type caller = null, [CallerMemberName] string callerName = "")
        {

            Log.Verbose($"{caller?.Name ?? _UNKNOWN_}::{callerName} --- {msg}");
        }



        private static int _pingDepth = 0;
        private const string _UNKNOWN_ = "???";

        private static string IndentBar() => string.Concat(Enumerable.Repeat("│  ", _pingDepth++));

        private static string IndentBarDecremented()
        {
            if (_pingDepth > 0) _pingDepth--;
            return string.Concat(Enumerable.Repeat("│  ", _pingDepth));
        }

        //public static void Ping(Type caller, [CallerMemberName] string callerMemberName = "")
        //{
        //    Log.Verbose($"{IndentBar()}-> Ping from {caller?.Name ?? _UNKNOWN_}::{callerMemberName}");
        //}
        // --- Ping Methods ---

        public static void Ping<T>([CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}");
        }

        public static void Ping<T>(string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Ping<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}");
        }

        public static void Ping<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Ping<T>(this T caller, Type[] genericParams, [CallerMemberName] string callerName = "")
        {
            string genericParamsString = $"<{string.Join(",", genericParams.Select(p => p.Name))}>";
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}{genericParamsString}");
        }

        // --- Pong Methods ---

        public static void Pong<T>([CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}");
        }

        public static void Pong<T>(string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Pong<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}");
        }

        public static void Pong<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Pong<T>(this T caller, Type[] genericParams, [CallerMemberName] string callerName = "")
        {
            string genericParamsString = $"<{string.Join(",", genericParams.Select(p => p.Name))}>";
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}{genericParamsString}");
        }

        public static void Pong<T>(this T caller, Type genericParam, [CallerMemberName] string callerName = "")
        {
            caller.Pong(new[] { genericParam }, callerName);
        }

        // --- PingPong Helpers ---

        public static void PingPong<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            caller.Ping(callerName);
            caller.Pong(callerName);
        }

        public static void PingPong<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            caller.Ping(msg, callerName);
            caller.Pong(msg, callerName);
        }
        public static TResult PingPong<T, TResult>(Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            return PingPong(typeof(T), action, msg, callerName);
        }
        public static TResult PingPong<T, TResult>(this T caller, Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            if (msg == null)
                caller.Ping(callerName);
            else
                caller.Ping(msg, callerName);

            var result = action();

            if (msg == null)
                caller.Pong(callerName);
            else
                caller.Pong(msg, callerName);

            return result;
        }

        public static TResult Pong<T, TResult>(this T caller, Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            var result = action();

            if (msg == null)
                caller.Pong(callerName);
            else
                caller.Pong(msg, callerName);

            return result;
        }
        /// <summary>
        /// Formats a sequence of strings into a multiline string with optional line numbering and prefix.
        /// </summary>
        /// <param name="lines">The sequence of strings to format.</param>
        /// <param name="prefix">Optional prefix to apply to each line.</param>
        /// <param name="includeIndex">If true, include line numbers starting at 1.</param>
        /// <returns>A single formatted string.</returns>
        public static string DumpString(this IEnumerable<string> lines, string prefix = "", bool includeIndex = false)
        {
            if (lines == null) return string.Empty;

            var sb = new StringBuilder();
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
    }
}
