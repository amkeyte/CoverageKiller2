using CoverageKiller2.Logging;
using Newtonsoft.Json;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Central manager for all Word application instances in the CoverageKiller2 system.
    /// Handles lifecycle and cleanup, including the special VSTO ThisAddIn instance.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0010
    /// </remarks>
    public class CKOffice_Word : IDisposable
    {
        private static CKOffice_Word _instance;

        /// <summary>
        /// Gets the singleton instance of <see cref="CKOffice_Word"/>.
        /// </summary>
        public static CKOffice_Word Instance
        {
            get
            {
                if (_instance == null) _instance = new CKOffice_Word();
                return _instance;
            }
        }

        private readonly List<CKApplication> _applications = new List<CKApplication>();
        private ThisAddIn _addinInstance;
        private bool _isRunning;
        private bool _disposedValue;

        public static bool IsTest => AppDomain.CurrentDomain.GetAssemblies()
                .Any(a => a.FullName.StartsWith("Microsoft.VisualStudio.TestPlatform")
                       || a.FullName.StartsWith("xunit")
                       || a.FullName.StartsWith("nunit"));

        static CKOffice_Word()
        {
            if (Debugger.IsAttached)
            {
                AppDomain.CurrentDomain.ProcessExit += (s, e) =>
                {
                    Log.Warning("Debugger stopped. Attempting emergency shutdown.");
                    try { Instance?.ShutDown(); } catch { }
                };

                AppDomain.CurrentDomain.DomainUnload += (s, e) =>
                {
                    Log.Warning("AppDomain unloading. Attempting emergency shutdown.");
                    try { Instance?.ShutDown(); } catch { }
                };
            }
        }

        // ... existing members


        private CKOffice_Word() { }

        /// <summary>
        /// Gets the registered application instances managed by this office context.
        /// </summary>
        public IEnumerable<CKApplication> Applications => _applications.Where(a => a != null);

        /// <summary>
        /// Indicates whether CKOffice_Word is currently running.
        /// </summary>
        public bool IsRunning => _isRunning;

        /// <summary>
        /// Gets the application instance representing the VSTO add-in, if registered.
        /// </summary>
        public CKApplication AddInApp => _applications.FirstOrDefault(a => a.IsAddIn);

        public static Tracer Tracer => new Tracer(typeof(CKOffice_Word), indentTabs: 20);

        /// <summary>
        /// Registers the VSTO add-in as a known guest instance. It will not be owned or disposed.
        /// </summary>
        /// <param name="addin">The add-in instance to register.</param>
        /// <returns>0 if registered successfully.</returns>
        public int TryPutAddin(ThisAddIn addin)
        {
            this.Ping();
            var wordApp = Globals.ThisAddIn.Application;
            if (_isTesting) return 0;
            _addinInstance = addin ?? throw new ArgumentNullException(nameof(addin));
            var addInApp = new CKApplication(addin.Application, default, isOwned: false);
            _applications.Add(addInApp);
            Log.Information("Registered ThisAddIn instance.");
            return 0;
        }

        /// <summary>
        /// Attempts to create a new owned CKApplication instance.
        /// </summary>
        /// <param name="app">The created application wrapper, or null if failed.</param>
        /// <param name="visible">Whether the Word UI should be visible.</param>
        /// <returns>The count of owned applications if successful; -1 if failed.</returns>
        /// <remarks>
        /// Version: CK2.00.00.0010
        /// </remarks>
        public int TryGetNewApp(out CKApplication app, bool visible = false)
        {
            this.Ping(msg: $"Found {_applications.Count} open CKApplication instances.");
            int pid = -1;

            try
            {
                var before = Process.GetProcessesByName("WINWORD").Select(p => p.Id).ToHashSet();
                Word.Application wordApp = new Word.Application { Visible = visible };
                System.Threading.Thread.Sleep(250);

                var after = Process.GetProcessesByName("WINWORD")
                                   .Where(p => !before.Contains(p.Id))
                                   .OrderByDescending(p => p.StartTime)
                                   .FirstOrDefault();

                if (after != null)
                    pid = after.Id;
                else
                    Log.Warning("Could not determine PID of new Word instance. Using -1.");

                app = new CKApplication(wordApp, pid, isOwned: true);
                _applications.Add(app);

                AppRecordManager.Add(app.PID);
                AppRecordManager.Save();

                Log.Information("New CKApplication({PID}) created and registered.", app.PID);
                this.Pong();
                return _applications.Count;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to create CKApplication:{PID} {Message}", pid, ex.Message);
                app = null;
                this.Pong();
                return -1;
            }
        }

        /// <summary>
        /// Starts CKOffice_Word and configures logging.
        /// Also cleans up orphaned Word processes from previous runs.
        /// </summary>
        /// <returns>0 if started; 1 if already running; -1 if error occurred.</returns>
        public int Start()
        {


            if (_isRunning)
            {
                Log.Warning("CKOffice_Word.Start called while already running. Start Refused.");
                return int.MinValue;
            }

            try
            {
                string logFile = LogTailLoader.GetLogFile();

                LoggingLoader.Configure(logFile, Serilog.Events.LogEventLevel.Debug);

                Log.Information("******************************************************************** CKOffice_Word started. ******************************************************************");

                _isRunning = true;

                Log.Information("Cleaning orphaned instances.");
                AppRecordManager.Load();
                AppRecordManager.CleanupOrphanedProcesses();

                return 0;
            }
            catch (Exception ex)
            {
                Log.Error("Error during CKOffice_Word startup: {Message}", ex.Message);
                return -1;
            }
        }

        private bool _crashing = false;
        private bool _isTesting;

        public void Crash(Exception ex, Type callerType, [MemberCallerName] string callerMember = null)
        {
            Log.Error($"Crashing {nameof(CKOffice_Word)}. Source: {callerType.Name}.{callerMember}");
            _crashing = true;
            if (ex != null)
            {
                Log.Error(ex, "Exception during crash from {Caller}.{Member}", callerType?.Name, callerMember);
            }

            for (int appIndex = 0; appIndex < Applications.Count(); appIndex++)
            {
                var app = Applications.ElementAt(appIndex);
                app.Crash(callerType, callerMember);
            }
            if (AddInApp != null)
            {
                AddInApp.Crash(callerType, callerMember);
                AddInApp.WordApp.Quit();
                AddInApp.Dispose();
            }

            ShutDown();
        }
        public int ShutDown()
        {
            this.Ping();

            if (!_isRunning)
            {
                Log.Warning("CKOffice_Word.ShutDown called but instance is not running.");
                return 1;
            }

            Log.Information("CKOffice_Word shutting down.");

            // Evaluate shutdown blocking only if not crashing
            bool blockShutDown = !_crashing && (AddInApp != null || _applications.Any(a => a.HasKeepOpenDocuments));

            foreach (var app in _applications.ToList())
            {
                if (!_crashing && (app == AddInApp || app.HasKeepOpenDocuments))
                {
                    Log.Information($"Application {app.PID} bypass shutting down.");
                    continue;
                }

                try
                {
                    app.Dispose();
                }
                catch (Exception ex)
                {
                    Log.Error("Error shutting down application: {Message}", ex.Message);
                }
            }

            if (!blockShutDown)
            {
                _applications.RemoveAll(a => a != AddInApp);

                try { LoggingLoader.Cleanup(); } catch { }

                try { LogTailLoader.Cleanup(); } catch { }

                _isRunning = false;
            }
            else
            {
                Log.Information("ThisAddIn still running. CKOffice_Word remains available.");
            }

            this.Pong();
            return 0;
        }

        protected virtual void Dispose(bool disposing)
        {
            this.Ping();
            if (!_disposedValue)
            {
                if (disposing)
                {
                    ShutDown();
                }
                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            this.Ping();
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

    /// <summary>
    /// Represents a record of a previously created Word application process.
    /// Used for crash recovery and cleanup of orphaned instances.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0009
    /// </remarks>
    public class AppRecord
    {
        /// <summary>
        /// The process ID of the Word instance.
        /// </summary>
        public int PID { get; set; }

        /// <summary>
        /// Optional tag for diagnostics or identification (e.g., "12345#ThisAddIn").
        /// </summary>
        public string Tag { get; set; }
    }

    /// <summary>
    /// Tracks known Word application process records for crash recovery.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0010
    /// </remarks>
    public static class AppRecordManager
    {
        private static readonly List<AppRecord> _records = new List<AppRecord>();

        /// <summary>
        /// Gets the current list of tracked AppRecords.
        /// </summary>
        public static IReadOnlyList<AppRecord> Records => _records.AsReadOnly();

        /// <summary>
        /// Adds a new AppRecord to the list.
        /// </summary>
        public static void Add(string pid, string tag = null)
        {
            var pid2 = int.Parse(pid);
            if (pid2 <= 0) return;
            _records.Add(new AppRecord { PID = pid2, Tag = tag });
            Save();
        }

        /// <summary>
        /// Saves the current list to Properties.Settings.Default as JSON.
        /// </summary>
        public static void Save()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_records, Formatting.Indented);
                Properties.Settings.Default.PreviousAppRecords = json;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Log.Error("Failed to save AppRecords: {Message}", ex.Message);
            }
        }

        /// <summary>
        /// Loads AppRecords from Properties.Settings.Default.
        /// </summary>
        public static void Load()
        {
            try
            {
                _records.Clear();
                var json = Properties.Settings.Default.PreviousAppRecords;
                if (!string.IsNullOrWhiteSpace(json))
                {
                    var list = JsonConvert.DeserializeObject<List<AppRecord>>(json);
                    if (list != null) _records.AddRange(list);
                }
            }
            catch (Exception ex)
            {
                Log.Warning("Could not load AppRecords from settings: {Message}", ex.Message);
            }
        }

        /// <summary>
        /// Kills orphaned WINWORD processes based on AppRecords.
        /// </summary>
        public static void CleanupOrphanedProcesses()
        {
            Log.Information("Checking for orphaned WINWORD processes...");

            var toRemove = new List<AppRecord>();

            foreach (var record in _records)
            {
                try
                {
                    var proc = Process.GetProcessById(record.PID);
                    if (proc.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
                    {
                        Log.Warning("Found orphaned WINWORD process (PID={PID}, Tag={Tag}). Attempting to terminate...", record.PID, record.Tag);
                        proc.Kill();
                        proc.WaitForExit(3000);
                        Log.Information("Terminated WINWORD process {PID}.", record.PID);
                        toRemove.Add(record);
                    }
                }
                catch (ArgumentException)
                {
                    toRemove.Add(record); // Process no longer exists
                }
                catch (Exception ex)
                {
                    Log.Warning("Could not terminate process {PID}: {Message}", record.PID, ex.Message);
                }
            }

            foreach (var rec in toRemove)
                _records.Remove(rec);

            Save();
        }
    }
}
