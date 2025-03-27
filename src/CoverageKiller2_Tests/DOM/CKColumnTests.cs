using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKColumnTests
    {
        //    // Constant for the test table number (one-based); adjust as needed.
        //    private const int TestTableNumber = 1;
        //    // Constant for the test column index (one-based); adjust as needed.
        //    private const int TestColumnIndex = 3;

        //    [TestMethod]
        //    public void CKColumn_Constructed_FromTableAndRectangularCellReference_HasCorrectIndexAndCount()
        //    {
        //        LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
        //        {
        //            // Ensure the document has at least one table.
        //            Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
        //                $"Test document must contain at least {TestTableNumber} table(s).");

        //            // Retrieve the test table.
        //            Word.Table wordTable = doc.Tables[TestTableNumber];
        //            // Create a CKTable from the Word table.
        //            CKTable ckTable = new CKTable(wordTable);

        //            // Get the total number of rows in the table.
        //            int rowCount = wordTable.Rows.Count;
        //            Assert.IsTrue(rowCount > 0, "The table must have at least one row.");

        //            // Create a rectangular cell reference for the entire column using a factory method.
        //            // This method is assumed to exist in CKCellRefRect.
        //            var colRef = CKCellRefRect.ForColumn(TestColumnIndex, rowCount);

        //            // Create the CKColumn from the table and the cell reference.
        //            CKColumn column = new CKColumn(ckTable, colRef);

        //            // Verify that the column index equals TestColumnIndex.
        //            Assert.AreEqual(TestColumnIndex, column.Index, "The CKColumn index should match the specified column index.");

        //            // Verify that the number of cells equals the row count.
        //            Assert.AreEqual(rowCount, column.Count, "The CKColumn should contain one cell per table row.");

        //            // Verify that the ToString() output contains the column index and cell count.
        //            string toString = column.ToString();
        //            Assert.IsTrue(toString.Contains($"Index: {TestColumnIndex}"), "ToString() should contain the column index.");
        //            Assert.IsTrue(toString.Contains($"Cells: {rowCount}"), "ToString() should contain the cell count.");
        //        });
        //    }
        //}
    }
}
