using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKTableGridTests
    {
        // Constant for the test table number; adjust if needed.
        private const int TestTableNumber = 1;

        [TestMethod]
        public void CKTableGrid_GetInstance_ReturnsSameInstanceForSameTableRange()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the test document has at least one table.
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                // Build a CKRange from the table's range.
                CKRange tableRange = new CKRange(wordTable.Range);
                // Get instance for this table range.
                CKTableGrid grid1 = CKTableGrid.GetInstance(wordTable);
                CKTableGrid grid2 = CKTableGrid.GetInstance(wordTable);

                Debug.WriteLine("Assertions:");
                Assert.IsNotNull(grid1, "GetInstance should return a valid CKTableGrid.");
                // For the same table range, GetInstance should return the same grid instance.
                Assert.AreSame(grid1, grid2, "GetInstance should return the same instance for the same table.");
                Debug.WriteLine(" OK");

            });
            Debug.WriteLine("New Test Starting");
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the test document has at least one table.
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable2 = doc.Tables[TestTableNumber];
                // Build a CKRange from the table's range.
                CKRange tableRange2 = new CKRange(wordTable2.Range);
                // Get instance for this table range.
                CKTableGrid grid12 = CKTableGrid.GetInstance(wordTable2);
                CKTableGrid grid22 = CKTableGrid.GetInstance(wordTable2);

                Debug.Write("Assertions:");
                Assert.IsNotNull(grid12, "GetInstance should return a valid CKTableGrid.");
                // For the same table range, GetInstance should return the same grid instance.
                Assert.AreSame(grid12, grid22, "GetInstance should return the same instance for the same table.");
                Debug.WriteLine(" OK");

            });


        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_ValidCoordinates_ReturnsGridCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                // Check that grid dimensions are set.
                Assert.IsTrue(grid.RowCount > 0, "RowCount should be greater than 0.");
                Assert.IsTrue(grid.ColCount > 0, "ColCount should be greater than 0.");

                // Using zero-based indices, get the top-left cell.
                GridCell cell = grid.GetGridCellAt(0, 0);
                Assert.IsNotNull(cell, "GetCellAt for valid coordinates should return a non-null GridCell.");
            });
        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_InvalidCoordinates_ReturnsNull()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                // Try invalid (negative) indices.
                GridCell cell = grid.GetGridCellAt(-1, -1);
                Assert.IsNull(cell, "GetCellAt with invalid indices should return null.");
            });
        }

        [TestMethod]
        public void CKTableGrid_GetRowCells_ReturnsValidCells()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                // Assume at least one row exists.
                int rowNumber = 1;
                var rowCells = grid.GetRowCells(rowNumber);
                Assert.IsNotNull(rowCells, "GetRowCells should return a valid collection.");
                Assert.IsTrue(rowCells.Any(), "GetRowCells should yield at least one cell.");
            });
        }

        [TestMethod]
        public void CKTableGrid_GetColumnCells_ReturnsValidCells()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                // Assume at least one column exists.
                int columnNumber = 1;
                var columnCells = grid.GetColumnCells(columnNumber);
                Assert.IsNotNull(columnCells, "GetColumnCells should return a valid collection.");
                Assert.IsTrue(columnCells.Any(), "GetColumnCells should yield at least one cell.");
            });
        }

        [TestMethod]
        public void CKTableGrid_Refresh_DoesNotThrowAndReturnsDiffs()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Test document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                // Create a CKTable from the Word table.
                CKTable ckTable = new CKTable(wordTable);

                // Call Refresh with getDiffs = true. We do not expect an exception.
                var diffs = CKTableGrid.Refresh(ckTable, true);
                // diffs can be null if there are no differences, so we just ensure the call succeeded.
                Assert.IsNotNull(diffs, "Refresh should return a diff collection (even if empty) when getDiffs is true.");
            });
        }
    }
}
