﻿using CoverageKiller2.DOM;
using Serilog;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Provides utilities for managing long-running document operations.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.03.0003
    /// </remarks>
    public static class LongOperationHelpers
    {
        public static bool PauseWithCountdown(string message = "Paused", int seconds = 30, bool allowCancel = false)
        {
            bool shouldContinue = false;
            bool userCanceled = false;

            var form = new Form
            {
                Text = "Pause for Inspection",
                Width = 400,
                Height = 160,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true
            };

            var label = new Label
            {
                Text = $"{message}\nContinuing automatically in {seconds} seconds...",
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter
            };

            var continueBtn = new Button
            {
                Text = "Continue Now",
                Dock = DockStyle.Left,
                Width = form.Width / (allowCancel ? 2 : 1)
            };

            continueBtn.Click += (s, e) =>
            {
                shouldContinue = true;
                form.Close();
            };

            Button cancelBtn = null;
            if (allowCancel)
            {
                cancelBtn = new Button
                {
                    Text = "Cancel",
                    Dock = DockStyle.Right,
                    Width = form.Width / 2
                };
                cancelBtn.Click += (s, e) =>
                {
                    userCanceled = true;
                    form.Close();
                };
            }

            var panel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
            panel.Controls.Add(continueBtn);
            if (allowCancel) panel.Controls.Add(cancelBtn);

            form.Controls.Add(label);
            form.Controls.Add(panel);

            var timer = new System.Windows.Forms.Timer { Interval = 1000 };
            timer.Tick += (s, e) =>
            {
                seconds--;
                label.Text = $"{message}\nContinuing automatically in {seconds} seconds...";
                if (seconds <= 0)
                {
                    shouldContinue = true;
                    form.Close();
                }
            };

            timer.Start();
            form.Show();

            // Manual event loop that lets Word GUI respond
            while (form.Visible)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(50);
            }

            timer.Stop();
            form.Dispose();

            return allowCancel ? shouldContinue && !userCanceled : true;
        }


        public static void DoStandardPause()
        {
            bool continueOperation = LongOperationHelpers.PauseWithCountdown(
                message: "Inspect and continue...",
                seconds: 30,
                allowCancel: true
            );


            if (!continueOperation)
            {
                Log.Warning("User canceled operation during pause checkpoint.");
                CKOffice_Word.Instance.ShutDown();
                return;
            }
            Log.Information("Continuing after pause");
        }
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

            var path = doc.FullPath;
            if (!File.Exists(path))
            {
                Log.Warning($"[TrySilentSave] File does not exist: {path}");
                return;
            }

            // Check if Word considers it readonly
            if (doc.ReadOnly)
            {
                Log.Warning($"[TrySilentSave] Word document is marked read-only. Falling back to backup save. {context}");
                SaveBackupCopy(doc, context);
                return;
            }

            // Check if file is writable on disk
            try
            {
                using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Write)) { }
            }
            catch (IOException)
            {
                Log.Warning($"[TrySilentSave] File is not writable. Falling back to backup save. {context}");
                SaveBackupCopy(doc, context);
                return;
            }

            try
            {
                doc.Saved = false;
                doc.Application.WordApp.ActiveDocument.Save();
                Log.Debug($"[TrySilentSave] Document saved successfully. {context}");
            }
            catch (Exception ex)
            {
                Log.Warning(ex, $"[TrySilentSave] Save failed unexpectedly. Falling back to backup. {context}");
                SaveBackupCopy(doc, context);
            }
        }

        private static void SaveBackupCopy(CKDocument doc, string context)
        {
            try
            {
                var dir = Path.GetDirectoryName(doc.FullPath);
                var baseName = Path.GetFileNameWithoutExtension(doc.FullPath);
                var ext = Path.GetExtension(doc.FullPath);

                int counter = 1;

                string backupPath = GetNextBackupPath(doc.FullPath);

                Log.Warning($"[TrySilentSave] Writing backup copy to: {backupPath}");

                doc.Application.WithSuppressedAlerts(() =>
                {
                    doc.SaveAs2(
                        backupPath,
                        fileFormat: Word.WdSaveFormat.wdFormatXMLDocument,
                        addToRecentFiles: false
                    );
                });

                Log.Information($"[TrySilentSave] Backup saved successfully: {backupPath}");
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"[TrySilentSave] Failed to write backup copy. {context}");
            }
        }
        private static string GetNextBackupPath(string originalPath)
        {
            var dir = Path.GetDirectoryName(originalPath);
            var fileName = Path.GetFileNameWithoutExtension(originalPath);
            var ext = Path.GetExtension(originalPath);

            // Remove any existing _Backup-XX suffix
            var baseName = System.Text.RegularExpressions.Regex.Replace(
                fileName,
                @"_Backup-\d{2}$", // match "_Backup-XX" at end
                ""
            );

            // Look for files that match the backup pattern for this baseName
            var existing = Directory.GetFiles(dir, $"{baseName}_Backup-??{ext}");

            int nextIndex = existing
                .Select(path => Path.GetFileNameWithoutExtension(path))
                .Select(name =>
                {
                    var suffix = name.Substring(name.LastIndexOf("-") + 1);
                    return int.TryParse(suffix, out var num) ? num : 0;
                })
                .DefaultIfEmpty(0)
                .Max() + 1;

            var backupPath = Path.Combine(dir, $"{baseName}_Backup-{nextIndex:00}{ext}");
            return backupPath;
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
