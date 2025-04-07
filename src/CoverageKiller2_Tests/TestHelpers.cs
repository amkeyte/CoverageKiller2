using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace CoverageKiller2.Test
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Runtime.InteropServices;
    using Word = Microsoft.Office.Interop.Word;
    [TestClass]
    public class WordSlayer
    {
        [TestMethod]
        public void ShowWordInstanceCount()
        {
            Console.WriteLine($"Number of open word applications found: {Process.GetProcessesByName("WINWORD").Length}");
        }
        [TestMethod]
        public void CloseAllWordInstances()
        {
            try
            {
                // Try to get a running instance of Word
                Word.Application wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");

                // Only close if there are no documents open (to avoid ruining someone's work)
                if (wordApp != null && wordApp.Documents.Count == 0)
                {
                    wordApp.Quit(false);
                    Marshal.ReleaseComObject(wordApp);
                    Console.WriteLine("Closed an idle Word instance.");
                }
                else
                {
                    Console.WriteLine("Word is busy. Not closing.");
                }
            }
            catch (COMException)
            {
                Console.WriteLine("No running Word instance found.");
            }
        }
        [TestMethod]
        public void KillAllWordProcesses()
        {
            foreach (var process in Process.GetProcessesByName("WINWORD"))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(); // Optional
                    Console.WriteLine("Process Killed.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to kill process {process.Id}: {ex.Message}");
                }
            }
        }
    }

    public static class TestHelpers
    {
        public static string DumpVisualRows(Base1JaggedList<Word.Cell> visualRows)
        {
            if (visualRows == null || visualRows.Count == 0)
                return "(no rows found)";

            var sb = new StringBuilder();

            foreach (var row in visualRows)
            {
                var rowText = row.Select(cell =>
                {
                    string text;
                    text = (cell == null) ? "NULL" :
                        CKTextHelper.Scrunch(cell?.Range.Text ?? "NULL");
                    var rowI = cell?.RowIndex.ToString() ?? "X";
                    var rowC = cell?.ColumnIndex.ToString() ?? "X";
                    var width = cell?.Width;
                    var height = cell?.Height;
                    return $"[{rowI},{rowC}] '{text}' [{width} x {height}]";
                });

                sb.AppendLine(string.Join(" | ", rowText));
            }

            return sb.ToString();
        }
        public static string DumpVisualRows(Base1JaggedList<GridCell> visualRows)
        {
            if (visualRows == null || visualRows.Count == 0)
                return "(no rows found)";

            var sb = new StringBuilder();

            foreach (var row in visualRows)
            {
                var rowText = row.Select(cell =>
                {
                    string text = cell?.COMCell?.Range?.Text ?? "NULL";
                    text = CKTextHelper.Scrunch(text);
                    var rowI = cell?.GridRow.ToString() ?? "X";
                    var rowC = cell?.GridCol.ToString() ?? "X";
                    var width = cell?.COMCell.Width;
                    var height = cell?.COMCell.Height;
                    return $"[{rowI},{rowC}] '{text}' [{width} x {height}]";
                });

                sb.AppendLine(string.Join(" | ", rowText));
            }

            return sb.ToString();
        }

        public static string DumpVisualRows(Word.Table table)
        {
            var allCells = table.Range.Cells.Cast<Word.Cell>().ToList();

            var grouped = allCells
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .Select(g =>
                {
                    var ordered = g.OrderBy(c => c.ColumnIndex);
                    return string.Join(" | ", ordered.Select(cell =>
                    {
                        var text = CKTextHelper.Scrunch(cell.Range.Text);
                        return $"[{cell.RowIndex},{cell.ColumnIndex}] '{text}'";
                    }));
                });

            return string.Join(Environment.NewLine, grouped);
        }

        public static List<string> DescribeTableRawCells(Word.Table table)
        {
            var lines = new List<string>();

            foreach (Word.Cell cell in table.Range.Cells)
            {
                string text = CKTextHelper.Scrunch(cell.Range.Text);
                lines.Add($"{text} ({cell.RowIndex},{cell.ColumnIndex})");
            }

            return lines;
        }

        /// <summary>
        /// Returns all CKTables in the document that contain merged cells.
        /// </summary>
        public static IReadOnlyList<CKTable> GetAllMergedTables(CKDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var mergedTables = new List<CKTable>();

            foreach (var ckTable in doc.Tables)
            {
                Word.Table wordTable = ckTable.COMTable;

                foreach (Word.Cell cell in wordTable.Range.Cells)
                {
                    if (IsMerged(cell))
                    {
                        mergedTables.Add(ckTable);
                        Debug.WriteLine($"Table: {mergedTables.Count} \"{ckTable.Cell(1).Text}\"");
                        break; // This table has at least one merged cell
                    }
                }
            }

            return mergedTables;
        }

        /// <summary>
        /// Returns the Nth CKTable in the document that contains merged cells.
        /// </summary>
        public static CKTable FindNthMergedTable(CKDocument doc, int index)
        {
            var mergedTables = GetAllMergedTables(doc);

            if (index < 1 || index > mergedTables.Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Document contains only {mergedTables.Count} merged tables.");

            return mergedTables[index - 1];
        }

        /// <summary>
        /// Determines whether a given Word cell is part of a merged region.
        /// </summary>
        private static bool IsMerged(Word.Cell cell)
        {
            try
            {
                // Merged cells usually span multiple logical cells
                return cell.Range.Cells.Count > 1;
            }
            catch
            {
                return false;
            }
        }
    }
}
