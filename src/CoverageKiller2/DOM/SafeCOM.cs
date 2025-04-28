using CoverageKiller2.DOM.Tables;
using Serilog;
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace CoverageKiller2.DOM
{

    /// <summary>
    /// Provides a helper to safely execute COM actions that may fail, marking the owning table dirty if needed.
    /// Supports optional retries and retry delay.
    /// </summary>
    public static class SafeCOM
    {

        /// <summary>
        /// Executes a COM operation safely. If a COMException occurs, the table is marked dirty.
        /// Retries are attempted with a small delay between tries.
        /// </summary>
        /// <param name="table">The owning table that will be marked dirty on failure.</param>
        /// <param name="action">The COM action to attempt.</param>
        /// <param name="maxRetries">Number of retries allowed if a COM failure occurs.</param>
        /// <param name="retryDelayMs">Delay (milliseconds) between retries. Default is 100ms.</param>
        /// <param name="rethrow">If true, rethrows the original exception after exhausting retries.</param>
        public static void Execute(CKTable table, Action action,
            int maxRetries = 1,
            int retryDelayMs = 100,
            bool rethrow = true,
            bool forceRefresh = false)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (maxRetries < 0) throw new ArgumentOutOfRangeException(nameof(maxRetries));
            if (retryDelayMs < 0) throw new ArgumentOutOfRangeException(nameof(retryDelayMs));

            int attempts = 0;

            while (true)
            {
                try
                {
                    action();
                    return; // Success! Exit
                }
                catch (COMException comEx)
                {
                    attempts++;
                    Log.Warning($"SafeCOMAction: COMException encountered (Attempt {attempts}): {comEx.Message}");

                    table.IsDirty = true;
                    if (forceRefresh) table.Refresh();

                    if (attempts > maxRetries)
                    {
                        if (rethrow)
                            throw;
                        else
                            return;
                    }

                    Log.Information($"SafeCOMAction: Retrying COM action (Attempt {attempts}/{maxRetries}) after {retryDelayMs}ms delay...");
                    Thread.Sleep(retryDelayMs);
                }
                catch (Exception ex)
                {
                    Log.Error($"SafeCOMAction: Unexpected exception encountered: {ex.Message}");

                    table.IsDirty = true;

                    if (rethrow)
                        throw;
                    else
                        return;
                }
            }
        }
    }
}


