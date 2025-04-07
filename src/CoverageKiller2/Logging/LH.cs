using Serilog;
using System;
using System.Runtime.CompilerServices;

namespace CoverageKiller2.Logging

{
    public static class LH
    {

        public static Exception LogThrow(Exception exception)
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
