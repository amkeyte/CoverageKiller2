using CoverageKiller2.DOM;
using Serilog;
using System;
using System.Runtime.InteropServices;

namespace CoverageKiller2.Helpers
{
    /// <summary>
    /// Provides utilities for managing long-running document operations.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.03.0003
    /// </remarks>
    public static class LongOperationHelpers
    {
        /// <summary>
        /// Attempts to save the document without prompting the user. Logs and suppresses errors.
        /// </summary>
        /// <param name="doc">The document to save.</param>
        /// <param name="context">Optional string describing the current operation for logging.</param>
        public static void TrySilentSave(CKDocument doc, string context = "")
        {
            if (doc == null)
            {
                Log.Warning("[TrySilentSave] No document provided.");
                return;
            }

            try
            {
                doc.Saved = false;
                doc.Application.WordApp.ActiveDocument.Save();
                Log.Information($"[TrySilentSave] Document saved successfully. {context}");
            }
            catch (COMException comEx)
            {
                Log.Warning($"[TrySilentSave] COMException during save. {context}: {comEx.Message}");
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"[TrySilentSave] Unexpected error during save. {context}");
            }
        }

        /// <summary>
        /// Tracks progress for long operations, logging percent complete and estimated time remaining.
        /// </summary>
        public class ProgressLogger
        {
            private readonly string _label;
            private readonly int _total;
            private readonly int _logEveryCount;
            private readonly TimeSpan _logEveryTime;
            private readonly DateTime _start;
            private DateTime _lastLogTime;
            private int _current = 0;

            public ProgressLogger(string label, int total, int logEveryCount = 50, double logEverySeconds = 2.0)
            {
                _label = label;
                _total = Math.Max(total, 1);
                _logEveryCount = Math.Max(logEveryCount, 1);
                _logEveryTime = TimeSpan.FromSeconds(logEverySeconds);
                _start = DateTime.UtcNow;
                _lastLogTime = _start;

                Log.Information($"[{_label}] Starting operation on {_total} items...");
            }

            /// <summary>
            /// Increments progress and logs status when thresholds are reached.
            /// </summary>
            public void Report()
            {
                _current++;
                var now = DateTime.UtcNow;
                var sinceLast = now - _lastLogTime;

                if (_current == _total || _current % _logEveryCount == 0 || sinceLast >= _logEveryTime)
                {
                    var elapsed = now - _start;
                    var percent = (double)_current / _total;
                    var estTotal = TimeSpan.FromTicks((long)(elapsed.Ticks / percent));
                    var remaining = estTotal - elapsed;

                    Log.Information($"[{_label}] {_current}/{_total} ({percent:P1}) complete. " +
                              $"Elapsed: {elapsed.TotalSeconds:n1}s. ETA: {remaining.TotalSeconds:n1}s.");

                    _lastLogTime = now;
                }
            }

            /// <summary>
            /// Logs final elapsed time at end of operation.
            /// </summary>
            public void Finish()
            {
                var elapsed = DateTime.UtcNow - _start;
                Log.Information($"[{_label}] Complete in {elapsed.TotalSeconds:n1}s.");
            }
        }
    }
}
