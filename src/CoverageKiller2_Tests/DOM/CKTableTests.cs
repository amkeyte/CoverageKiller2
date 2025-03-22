using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKTableTests
    {
        //        The error "The RPC server is unavailable" (HRESULT: 0x800706BA) typically indicates that the COM infrastructure can't communicate with the Word process. In the context of Word automation (or VSTO), this error may be caused by one or more of the following issues:

        //Word Process Issues:
        //The Word process may have crashed, terminated unexpectedly, or not started at all.If Word is not running or has become unresponsive, COM calls will fail.
        //        COM/DCOM Configuration:
        //Misconfigured DCOM security settings or insufficient permissions can prevent your process from connecting to the Word COM server.This includes both local and remote DCOM configurations.
        //        Network/Firewall Restrictions:
        //Although typically Word automation runs locally, if there are network or firewall settings interfering with RPC communication, you can see this error.

        //Bitness Mismatch:
        //A mismatch between the bitness (32-bit vs. 64-bit) of your test runner and the installed version of Office can cause communication issues.

        //Premature COM Object Release:
        //If a COM object has been released (or garbage collected) before all necessary calls are made, subsequent calls may attempt to use an invalid reference, triggering the error.

        //Reviewing these areas should help you pinpoint the root cause in your specific scenario.



        [TestMethod]
        public void CKTable_Constructor_LoadsTableSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the document has at least one table.
                Assert.IsTrue(doc.Tables.Count > 0, "Test document must contain at least one table.");

                // Get the first table (Word collections are 1-based).
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);

                // Verify that the underlying COMTable property is set correctly.
                Assert.IsNotNull(ckTable.COMTable, "COMTable property should not be null.");
                Assert.AreEqual(wordTable.Range.Start, ckTable.COMTable.Range.Start,
                    "The range start of the underlying COMTable should match.");
            });
        }


        [TestMethod]
        public void CKTable_Cell_MatchesDirectWordCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the document contains at least one table.
                Assert.IsTrue(doc.Tables.Count > 0, "Test document must contain at least one table.");

                // Retrieve the first table.
                Word.Table wordTable = doc.Tables[1];

                // Create a CKTable from the Word table.
                CKTable ckTable = new CKTable(wordTable);

                // Create a CKCellReference for cell (1,1) using the CKTable-based constructor.
                // Retrieve the CKCell using our CKTable.Cell() method.
                CKCell ckCell = ckTable.Cell(CKCellRefRect.ForCell(1, 1));

                // Directly retrieve the Word cell for (1,1) from the Word table and wrap it.
                Word.Cell wordCell = wordTable.Cell(1, 1);
                CKCell directCell = new CKCell(wordCell);



                // Use the new equality method to compare the CKCell wrappers.
                Assert.AreEqual(ckCell, directCell,
                    "The CKCell from CKTable.Cell(x,y) should equal the COM Table.Cell(x,y).");
            });
        }

        // Constant for the test table number; update this if needed.
        private const int TestTableNumber = 12;

        /// <summary>
        /// Verifies that a single cell reference can be created using the ForCell factory method on the test table.
        /// </summary>
        [TestMethod]
        public void CKTable_FactoryMethods_SingleCellInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Test document must contain at least {TestTableNumber} tables.");

                // Use the specified test table.
                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                // Create a cell reference for the cell at (1,1) using the rectangular factory method.
                var cellRef = CKCellRefRect.ForCell(1, 1);

                // Retrieve the cell via CKTable.
                CKCell cell = ckTable.Cell(cellRef);
                Assert.IsNotNull(cell, "A single cell reference should return a valid CKCell.");
            });
        }

        /// <summary>
        /// Verifies that a row reference can be created using the ForRow factory method on the test table.
        /// </summary>
        [TestMethod]
        public void CKTable_FactoryMethods_RowReferenceInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Test document must contain at least {TestTableNumber} tables.");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                int colCount = wordTable.Columns.Count;
                // Create a cell reference covering the entire row 2.
                var rowRef = CKCellRefRect.ForRow(2, colCount);

                // Retrieve the cell (typically the first cell in that row).
                CKCell cell = ckTable.Cell(rowRef);
                Assert.IsNotNull(cell, "A row cell reference should return a valid CKCell.");
            });
        }

        /// <summary>
        /// Verifies that a column reference can be created using the ForColumn factory method on the test table.
        /// </summary>
        [TestMethod]
        public void CKTable_FactoryMethods_ColumnReferenceInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Test document must contain at least {TestTableNumber} tables.");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                int rowCount = wordTable.Rows.Count;
                // Create a cell reference covering the entire column 3.
                var colRef = CKCellRefRect.ForColumn(3, rowCount);

                // Retrieve a cell from that column.
                CKCell cell = ckTable.Cell(colRef);
                Assert.IsNotNull(cell, "A column cell reference should return a valid CKCell.");
            });
        }

        /// <summary>
        /// Verifies that a rectangular cell reference can be created using the ForRectangle factory method on the test table.
        /// </summary>
        [TestMethod]
        public void CKTable_FactoryMethods_RectangularReferenceInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber, $"Test document must contain at least {TestTableNumber} tables.");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                // Choose coordinates that are within the bounds of the table.
                // For example, assume the test table has at least 5 rows and 6 columns.
                var rectRef = CKCellRefRect.ForRectangle(2, 3, 4, 5);

                // Retrieve the cell from the rectangular reference.
                CKCell cell = ckTable.Cell(rectRef);
                Assert.IsNotNull(cell, "A rectangular cell reference should return a valid CKCell.");
            });
        }
    }
}
