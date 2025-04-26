using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2._TestOperators
{
    [TestClass]
    public class MyWordExplorerClass
    {
        //******* Standard Rigging ********
        public TestContext TestContext { get; set; }
        private string _testFilePath;
        private CKDocument _testFile;

        [TestInitialize]
        public void Setup()
        {
            Log.Information($"Running test => {GetType().Name}::{TestContext.TestName}");
            RandomTestHarness.PreserveTempFilesAfterTest = true;

            _testFilePath = RandomTestHarness.TestFile1;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath, visible: true);
        }
        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: true);
            Log.Information($"Completed test => {GetType().Name}::{TestContext.TestName}; status: {TestContext.CurrentTestOutcome}");
        }
        //******* End Standard Rigging ********
        [TestMethod]
        public void SeeWhatWordDoes5()
        {
            var workspace = _testFile.Application.GetShadowWorkspace(true);
            workspace.CloneFrom(_testFile.Tables[1]);
            workspace.ShowDebuggerWindow();
        }

        [TestMethod]
        public void ExploreSetWidthEffects()
        {
            try
            {

                var workspace = _testFile.Application.GetShadowWorkspace(visible: true, keepAlive: true);

                // Step 1: Pick the source table (first table for now)
                var ckTable = _testFile.Tables[1];
                var sourceTable = ckTable.COMTable;
                Assert.IsNotNull(sourceTable, "No table found in document.");

                // Step 2: Set up test parameters
                var testWidth = 100f; // in points (100 pt ≈ 1.4 inch)
                var styles = new[]
                {
                Word.WdRulerStyle.wdAdjustNone,
                Word.WdRulerStyle.wdAdjustFirstColumn,
                Word.WdRulerStyle.wdAdjustProportional
            };

                // Step 3: Clone multiple copies of the table and apply settings
                foreach (var style in styles)
                {
                    // Clone table into workspace
                    var newTable = workspace.CloneFrom(ckTable);

                    // Apply SetWidth to each column in the new table
                    foreach (Word.Cell cell in newTable.Columns[2])
                    {
                        cell.SetWidth(testWidth, style);
                    }

                    // Insert a paragraph after each table to make it easier to view
                    var paragraphAfter = newTable.COMRange.Next();
                    paragraphAfter.InsertParagraphAfter();
                    paragraphAfter.Text = $"--- Applied SetWidth({testWidth}, {style}) ---";
                }

                workspace.ShowDebuggerWindow();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Exception occured during test.");
            }


        }
    }
}


//public static List<List<Word.Cell>> GetVisualRows(Word.Table table)
//        {
//            return table.Range.Cells
//                .Cast<Word.Cell>()
//                .GroupBy(c => c.RowIndex)
//                .OrderBy(g => g.Key)
//                .Select(g => g.OrderBy(c => c.ColumnIndex).ToList())
//                .ToList();
//        }


//        void RejoinTables(Word.Table top, Word.Table bottom)
//        {
//            if (top == null || bottom == null)
//                throw new ArgumentNullException("Both top and bottom tables must be non-null.");

//            if (bottom.Range.Start <= top.Range.End)
//                throw new ArgumentException("Bottom table must be after top table.");

//            // Get the range between the two tables — likely a single paragraph.
//            var gapRange = top.Range.Document.Range(top.Range.End, bottom.Range.Start);
//            gapRange.Delete(); // Yeet the paragraph Word shoved in there after Split()

//            //// Move rows from bottom to top
//            //while (bottom.Rows.Count > 0)
//            //{
//            //    Word.Row row = bottom.Rows[1];
//            //    row.Range.Cut();
//            //    top.Rows.Add();
//            //    top.Rows[top.Rows.Count].Range.Paste();
//            //}

//            // Kill the empty shell of a table we just pillaged
//            //bottom.Delete();
//        }
//        public static void SplitByParagraphHack(Word.Cell lastTopCell)
//        {
//            if (lastTopCell == null)
//                throw new ArgumentNullException(nameof(lastTopCell));

//            var splitPoint = lastTopCell.Range.Duplicate;
//            splitPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
//            splitPoint.InsertParagraphAfter();
//        }

//        [TestMethod]
//        public void SeeWhatWordDoes4()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var tIndex = 2;
//                var table = doc.Tables[tIndex].COMTable;
//                var wDoc = doc.COMDocument;
//                string.Join(" | ", TestHelpers.DescribeTableRawCells(table));

//                var x = TableGridCrawler2.GetVisualRows(table);
//                Debug.WriteLine("Visual rows");
//                Debug.WriteLine(TestHelpers.DumpVisualRows(x));
//                Debug.WriteLine("normal");
//                TableGridCrawler2.NormalizeJaggedList(x);
//                Debug.WriteLine(TestHelpers.DumpVisualRows(x));

//                var xx = table.Cell(1, 1).Row.Height;
//                Debug.WriteLine(xx);

//                //table 1 "A\r\aB\r\aC\r\a\r\aD\r\aE\r\aF\r\a\r\aG\r\aH\r\aI\r\a\r\aJ\r\aK\r\aL\r\a\r\a"
//                //table 2 ABDE\r\n\aC\r\n\a\r\n\a\r\n\aF\r\n\a\r\n\aG\r\n\aH\r\n\aI\r\n\a\r\n\aJ\r\n\aK\r\n\aL\r\n\a\r\n\a"
//                //table 2 ABDE\r\n\aC\r\n\a\r\n\a\r\n\aF\r\n\a\r\n\aG\r\n\aH\r\n\aI\r\n\a\r\n\aJ\r\n\aK\r\n\aL\r\n\a\r\n\a"
//                //table 3 A\r\aB\r\aC\r\a\r\aDEF\r\a\r\aG\r\aH\r\aI\r\a\r\a"
//                //table 4 ADG\r\aB\r\aC\r\a\r\a\r\aE\r\aF\r\a\r\a\r\aH\r\aI\r\a\r\a"
//                //xx = "ADG\r\a\r\a\r\a\r\a\r\a\r\a\r\a\r\a\r\aH\r\aI\r\a\r\a"


//                //var xx = table.Cell(1, 1);
//                //var xy = table.Cell(2, 1);
//                //Debug.WriteLine($"xx = {xx.Range.Text}; xy = {xy.Range.Text}");

//            });
//        }

//        [TestMethod]
//        public void SeeWhatWordDoes3()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var tIndex = 1;
//                var table = doc.Tables[tIndex].COMTable;
//                var wDoc = doc.COMDocument;
//                try
//                {
//                    var row = 4;
//                    var topIndex = tIndex;
//                    var bottomIndex = topIndex + 1;
//                    var lastTop = table.Cell(3, 3);
//                    SplitByParagraphHack(lastTop);
//                    //table.Split(row);
//                    var top = wDoc.Tables[topIndex];
//                    var bottom = wDoc.Tables[bottomIndex];


//                    var x = TestHelpers.DescribeTableRawCells(top);
//                    var y = TestHelpers.DescribeTableRawCells(bottom);





//                    try
//                    {
//                        Debug.WriteLine(string.Join(" | ", TestHelpers.DescribeTableRawCells(top)));

//                    }
//                    catch (Exception ex)
//                    {
//                        if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.ObjectDeleted)) throw;

//                        Debug.WriteLine("Table Deleted");
//                    }
//                    Debug.WriteLine("\nsplit\n");
//                    try
//                    {
//                        Debug.WriteLine(string.Join(" | ", y));

//                    }
//                    catch (Exception ex)
//                    {
//                        if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.ObjectDeleted)) throw;

//                        Debug.WriteLine("Table Deleted");
//                    }

//                    RejoinTables(top, bottom);
//                    var z = TestHelpers.DescribeTableRawCells(top);

//                    Debug.WriteLine("\nRejoined\n");
//                    try
//                    {
//                        Debug.WriteLine(string.Join(" | ", z));

//                    }
//                    catch (Exception ex)
//                    {
//                        if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.ObjectDeleted)) throw;

//                        Debug.WriteLine("oops");
//                    }


//                }
//                catch (Exception ex)
//                {
//                    if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.ObjectDeleted)) throw;

//                    Debug.WriteLine("Table problem");
//                }
//            });
//        }
//        [TestMethod]
//        public void SeeWhatWordDoes2()
//        {

//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[1];
//                var table2 = table.COMTable;
//                var sb = new StringBuilder();
//                var cells = table.COMRange.Cells.ToList();
//                //var chars = table.COMRange.Characters;
//                try
//                {

//                    //full table
//                    var x = table2.Columns[1];
//                    var cellText = string.Join(", ", x.Cells.ToList().Select(c => $"[{c.Range.Text}]"));
//                    Debug.WriteLine($"table columns good {cellText}");
//                }
//                catch (Exception ex)
//                {
//                    if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.MixedCellWidths)) throw;

//                    Debug.WriteLine("No columns for table");
//                }
//                try
//                {
//                    var lastCell = cells.Last();
//                    var x2 = lastCell.Range.Columns[1];
//                    var cellText = string.Join(", ", x2.Cells.ToList().Select(c => $"[{c.Range.Text}]"));

//                    Debug.WriteLine($"range columns good. {cellText}");

//                }
//                catch (Exception ex)
//                {
//                    if (!ExceptionDetail.Is(ex, KnownExceptions.VSTO.MixedCellWidths)) throw;

//                    Debug.WriteLine("No columns for last cell range.");
//                }




//            });
//        }


//        [TestMethod]
//        public void SeeWhatWordDoes()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[2];
//                var sb = new StringBuilder();
//                var cells = table.COMRange.Cells.ToList();

//                foreach (var cell in cells)
//                {
//                    var cellText = CKTextHelper.Scrunch(cell.Range.Text);
//                    int wdIndex = cells.IndexOf(cell);
//                    sb.AppendLine($"Cell: {wdIndex} | coord:{cell.RowIndex},{cell.ColumnIndex} | '{cellText}' ");

//                    string upResult = TryDirection(table, cells, cell, -1, 0, "Up");
//                    string downResult = TryDirection(table, cells, cell, 2, 0, "Down");
//                    string leftResult = TryDirection(table, cells, cell, 0, -1, "Left");
//                    string rightResult = TryDirection(table, cells, cell, 0, 1, "Right");

//                    sb.AppendLine(upResult);
//                    sb.AppendLine(downResult);
//                    sb.AppendLine(leftResult);
//                    sb.AppendLine(rightResult);
//                }

//                Debug.Print(sb.ToString());
//            });
//        }

//        private string TryDirection(CKTable table, System.Collections.Generic.List<Microsoft.Office.Interop.Word.Cell> cells, Microsoft.Office.Interop.Word.Cell origin, int rowOffset, int colOffset, string label)
//        {
//            try
//            {
//                _ = table.COMTable.Cell(origin.RowIndex, origin.ColumnIndex);
//                var targetCell = table.COMTable.Cell(origin.RowIndex + rowOffset, origin.ColumnIndex + colOffset);
//                targetCell = cells.FirstOrDefault(c => c.Range.COMEquals(targetCell.Range));

//                if (targetCell == null || targetCell.Range.COMEquals(origin.Range))
//                    throw new System.Runtime.InteropServices.COMException($"At {label} edge of table.");

//                var text = CKTextHelper.Scrunch(targetCell.Range.Text);
//                int index = cells.IndexOf(targetCell);
//                return $"\t {label}: {index} | coord:{targetCell.RowIndex},{targetCell.ColumnIndex} | '{text}' ";
//            }
//            catch (System.Runtime.InteropServices.COMException ex)
//            {
//                return $"\t {label}: Exception: {ex.Message}";
//            }
//        }
//    }
//}
