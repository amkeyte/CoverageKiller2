using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKRowTests
    {
        // Constant for the test table number. Adjust as necessary.
        private const int TestTableNumber = 12;
        // Constant for the row index to test (one-based).
        private const int TestRowIndex = 2;

        [TestMethod]
        public void CKRow_Constructor_LoadsRowSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the document has at least TestTableNumber tables.
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Document must contain at least {TestTableNumber} tables.");

                // Use the specified test table.
                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                // Get the total number of columns in the table.
                int columnCount = wordTable.Columns.Count;
                // Create a rectangular cell reference for an entire row.
                var cellRef = CKCellRefRect.ForRow(TestRowIndex, columnCount);

                // Create a CKRow instance.
                CKRow row = new CKRow(ckTable, cellRef);

                // Verify that the CKRow is constructed.
                Assert.IsNotNull(row, "CKRow should be constructed successfully.");
                // Verify that the row's Index property returns the expected row number.
                Assert.AreEqual(TestRowIndex, row.Index, "Row index should match the provided row index.");
            });
        }

        [TestMethod]
        public void CKRow_CellsEnumeration_YieldsCorrectCellCount()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Document must contain at least {TestTableNumber} tables.");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);
                int columnCount = wordTable.Columns.Count;
                var cellRef = CKCellRefRect.ForRow(TestRowIndex, columnCount);
                CKRow row = new CKRow(ckTable, cellRef);

                // Verify that the number of cells enumerated equals the number of columns.
                int expectedCellCount = columnCount;
                int actualCount = row.Cells.Count;
                Assert.AreEqual(expectedCellCount, actualCount, "The row should contain as many cells as there are columns.");
            });
        }
    }
}
