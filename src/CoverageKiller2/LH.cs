using Serilog;
using System;

namespace CoverageKiller2
{
    internal static class LH
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
            Result
        }
        public static string TraceCaller(params object[] paramPairs)
        {

            // Validate pairs length
            if (paramPairs.Length % 2 != 0)
            {
                throw new ArgumentException("The number of parameters must be even, as they come in pairs.");
            }

            PP? traceType = null;
            if (paramPairs[0] is null
                || paramPairs[0].GetType() != typeof(LH.PP))
            {
                throw new ArgumentException("Bad caller typeflag at param1.");
            }
            else
            {
                traceType = (PP)paramPairs[0];
            }

            string message = paramPairs[1] as string ?? string.Empty;

            string formattedPairs = "";

            switch (traceType)
            {
                case PP.Enter:
                    {
                        message = message == string.Empty ?
                            "Entering member:\n" : $"{message}\n";
                        string className = paramPairs[2].ToString();
                        string methodName = paramPairs[3].ToString();

                        for (int i = 4; i < paramPairs.Length; i += 2)
                        {
                            // Assuming pairs[i] is the name and pairs[i + 1] is the value
                            string name = paramPairs[i].ToString();
                            string value = paramPairs[i + 1].ToString();
                            formattedPairs += $"[{name} = {value}]";
                        }

                        return $"TRACE => {className}.{methodName} :: {message}\n\t{formattedPairs}";

                    }
                case PP.Result:
                    {
                        message = message == string.Empty ?
                                                    "Entering member:\n" : $"{message}\n";
                        string className = paramPairs[2].ToString();
                        string methodName = paramPairs[3].ToString();

                        for (int i = 4; i < paramPairs.Length; i += 2)
                        {
                            // Assuming pairs[i] is the name and pairs[i + 1] is the value
                            string name = paramPairs[i].ToString();
                            string value = paramPairs[i + 1].ToString();
                            formattedPairs += $"[{name} = {value}]";
                        }

                        return $"TRACE => {className}.{methodName} :: {message}\n\t{formattedPairs}";
                    }
                default:
                    throw new ArgumentException("Bad TraceType");

            }

        }
    }
}
