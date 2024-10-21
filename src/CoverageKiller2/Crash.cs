using Serilog;
using System;

namespace CoverageKiller2
{
    internal static class Crash
    {

        public static Exception LogThrow(Exception exception)
        {
            Log.Error(exception, exception.Message);
            return exception;
        }
    }
}
