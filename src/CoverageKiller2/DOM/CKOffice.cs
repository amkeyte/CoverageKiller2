using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Central manager for all Word application instances in the CoverageKiller2 system.
    /// Handles lifecycle and cleanup, including the special VSTO ThisAddIn instance.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0005
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

        /// <summary>
        /// Registers the VSTO add-in as a known guest instance. It will not be owned or disposed.
        /// </summary>
        /// <param name="addin">The add-in instance to register.</param>
        /// <returns>0 if registered successfully.</returns>
        public int TryPutAddin(ThisAddIn addin)
        {
            if (addin == null) throw new ArgumentNullException(nameof(addin));
            _addinInstance = addin;
            var addInApp = new CKApplication(addin.Application);
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
        public int TryGetNewApp(out CKApplication app, bool visible = false)
        {
            try
            {
                Word.Application wordApp = new Word.Application { Visible = visible };
                app = new CKApplication(wordApp);
                _applications.Add(app);
                Log.Information("New CKApplication created and registered.");
                return _applications.Count;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to create CKApplication: {Message}", ex.Message);
                app = null;
                return -1;
            }
        }

        /// <summary>
        /// Starts CKOffice_Word and configures logging.
        /// Safe to call multiple times; subsequent calls are no-ops.
        /// </summary>
        /// <returns>0 if started; 1 if already running; -1 if error occurred.</returns>
        public int Start()
        {
            if (_isRunning)
            {
                Log.Information("CKOffice_Word.Start called while already running. No action taken.");
                return 1;
            }

            try
            {
                string logFile = LogTailLoader.GetLogFile();
                LoggingLoader.Configure(logFile, Serilog.Events.LogEventLevel.Debug);
                Log.Information("CKOffice_Word started.");
                _isRunning = true;
                return 0;
            }
            catch (Exception ex)
            {
                Log.Error("Error during CKOffice_Word startup: {Message}", ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// Shuts down all owned applications and cleans up logging.
        /// If ThisAddIn is still running, CKOffice_Word stays active.
        /// </summary>
        /// <returns>0 if shutdown was performed; 1 if not needed.</returns>
        public int ShutDown()
        {
            if (!_isRunning)
            {
                Log.Warning("CKOffice_Word.ShutDown called but instance is not running.");
                return 1;
            }

            Log.Information("CKOffice_Word shutting down.");

            bool hasAddIn = AddInApp != null;

            foreach (var app in _applications.ToList())
            {
                if (app == AddInApp) continue;

                try
                {
                    app.Dispose();
                }
                catch (Exception ex)
                {
                    Log.Error("Error shutting down application: {Message}", ex.Message);
                }
            }

            _applications.RemoveAll(a => a != AddInApp);

            try { LoggingLoader.Cleanup(); } catch { }
            try { LogTailLoader.Cleanup(); } catch { }

            if (hasAddIn)
            {
                Log.Information("ThisAddIn still running. CKOffice_Word remains available.");
                return 0;
            }

            _isRunning = false;
            return 0;
        }

        /// <summary>
        /// Disposes the CKOffice_Word singleton.
        /// </summary>
        /// <param name="disposing">True when called directly, false during finalization.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    ShutDown();
                }

                _disposedValue = true;
            }
        }

        /// <summary>
        /// Disposes the singleton and performs shutdown logic.
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
