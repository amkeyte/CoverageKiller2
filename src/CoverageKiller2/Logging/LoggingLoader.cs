using Serilog;
using Serilog.Events;

namespace CoverageKiller2.Logging
{

    /// <summary>
    /// stuff goes here
    /// </summary>
    public class LoggingLoader
    {
        /// <summary>
        /// Gets the current log event level.
        /// </summary>
        public static LogEventLevel Level { get; private set; }
        public static string LogFile { get; private set; }
        /// <summary>
        /// Configures Serilog with the specified log file and log level.
        /// </summary>
        /// <param name="logFile">The file to which log events will be written.</param>
        /// <param name="logEventLevel">The minimum log event level.</param>
        public static void Configure(string logFile, LogEventLevel logEventLevel)
        {
            LogFile = logFile;
            Level = logEventLevel;

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is(logEventLevel)    // Set the minimum log level
                .WriteTo.File(logFile)
                //.WriteTo.Async(a => a.File(logFile)) // Log to a file asynchronously
                .CreateLogger();
        }

        /// <summary>
        /// Cleans up the logger, ensuring all log events are flushed before closing.
        /// </summary>
        public static void Cleanup()
        {
            Log.CloseAndFlush(); // Flush all log events and close the logger
        }

    }

}
