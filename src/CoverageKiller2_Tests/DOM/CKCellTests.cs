using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKCellTests
    {
        /// <summary>
        /// Verifies that the CKCell constructor properly wraps a Word.Cell COM object using the new reference system.
        /// </summary>
        [TestMethod]
        public void CKCell_Constructor_WrapsCOMCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                Word.Table table = doc.Tables[1];
                CKTable ckTable = new CKTable(table);
                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell ckCell = ckTable.Converters.GetCell(ckTable, cellRef);

                Assert.IsNotNull(ckCell.COMCell, "CKCell should wrap a valid COMCell.");
            });
        }

        /// <summary>
        /// Verifies that the Text property of CKCell gets and sets the cell text correctly.
        /// </summary>
        [TestMethod]
        public void CKCell_TextProperty_GetSet_Works()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                Word.Table table = doc.Tables[1];
                CKTable ckTable = new CKTable(table);
                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell ckCell = ckTable.Converters.GetCell(ckTable, cellRef);

                string originalText = ckCell.Text;
                string newText = originalText + " Test";

                ckCell.Text = newText;
                Assert.IsTrue(CKTextHelper.ScrunchEquals(newText, ckCell.Text), "The Text property should update to the new value.");

                ckCell.Text = originalText;
            });
        }

        /// <summary>
        /// Verifies that the BackgroundColor property of CKCell gets and sets the background color correctly.
        /// </summary>
        [TestMethod]
        public void CKCell_BackgroundColor_GetSet_Works()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                Word.Table table = doc.Tables[1];
                CKTable ckTable = new CKTable(table);
                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell ckCell = ckTable.Converters.GetCell(ckTable, cellRef);

                Word.WdColor originalColor = ckCell.BackgroundColor;

                ckCell.BackgroundColor = Word.WdColor.wdColorRed;
                Assert.AreEqual(Word.WdColor.wdColorRed, ckCell.BackgroundColor, "Background color should update to wdColorRed.");

                ckCell.BackgroundColor = originalColor;
            });
        }

        /// <summary>
        /// Verifies that the ForegroundColor property of CKCell gets and sets the foreground color correctly.
        /// </summary>
        [TestMethod]
        public void CKCell_ForegroundColor_GetSet_Works()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                Word.Table table = doc.Tables[1];
                CKTable ckTable = new CKTable(table);
                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell ckCell = ckTable.Converters.GetCell(ckTable, cellRef);

                Word.WdColor originalColor = ckCell.ForegroundColor;

                ckCell.ForegroundColor = Word.WdColor.wdColorBlue;
                Assert.AreEqual(Word.WdColor.wdColorBlue, ckCell.ForegroundColor, "Foreground color should update to wdColorBlue.");

                ckCell.ForegroundColor = originalColor;
            });
        }
    }
}
