using Serilog;
using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;

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
        public CKDebugException(string message) : base(message) { }

        /// <inheritdoc/>
        public CKDebugException(string message, Exception innerException)
            : base(message, innerException) { }

        /// <inheritdoc/>
        protected CKDebugException(SerializationInfo info, StreamingContext context)
                : base(info, context)
        {
        }
    }
    public static class LH
    {

        public static Exception LogThrow(Exception exception = null, [CallerMemberName] string callerName = "")
        {
            Log.Error(exception, exception.Message);
            return exception;
        }

        /// <summary>
        /// Logs an exception using Serilog and optionally rethrows it.
        /// </summary>
        /// <param name="ex">The exception to log.</param>
        /// <param name="context">Optional context message to include with the log.</param>
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
        public static void Checkpoint(string message, Type caller = null, [CallerMemberName] string callerName = "")
        {

            Log.Verbose($"{caller?.Name ?? _UNKNOWN_}::{callerName} --- {message}");
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

        public static void Ping<T>(string message, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {message}");
        }

        public static void Ping<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}");
        }

        public static void Ping<T>(this T caller, string message, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {message}");
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

        public static void Pong<T>(string message, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {message}");
        }

        public static void Pong<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}");
        }

        public static void Pong<T>(this T caller, string message, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {message}");
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

        public static void PingPong<T>(this T caller, string message, [CallerMemberName] string callerName = "")
        {
            caller.Ping(message, callerName);
            caller.Pong(message, callerName);
        }

        public static TResult PingPong<T, TResult>(this T caller, Func<TResult> action, string message = null, [CallerMemberName] string callerName = "")
        {
            if (message == null)
                caller.Ping(callerName);
            else
                caller.Ping(message, callerName);

            var result = action();

            if (message == null)
                caller.Pong(callerName);
            else
                caller.Pong(message, callerName);

            return result;
        }

        public static TResult Pong<T, TResult>(this T caller, Func<TResult> action, string message = null, [CallerMemberName] string callerName = "")
        {
            var result = action();

            if (message == null)
                caller.Pong(callerName);
            else
                caller.Pong(message, callerName);

            return result;
        }

    }
}
