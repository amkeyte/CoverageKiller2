using Serilog;
using Serilog.Events;

public class LoggingLoader
{
    private static ILogger _logger;
    public static LogEventLevel Level { get; private set; }
    // Method to configure Serilog
    public static void Configure(string logFile, LogEventLevel logEventLevel)
    {
        Level = logEventLevel;

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Is(logEventLevel)
            .WriteTo.Async(a => a.File(logFile)) // Log to a file asynchronously
            .CreateLogger();
    }

    // Clean up method for the logger
    public static void Cleanup()
    {
        Log.CloseAndFlush(); // Ensures that all log events are flushed before closing
    }
}
