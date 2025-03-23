using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKCellTests
    {
        /// <summary>
        /// Verifies that the CKCell constructor properly wraps a Word.Cell COM object.
        /// </summary>
        [TestMethod]
        public void CKCell_Constructor_WrapsCOMCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the document has at least one table.
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                // Retrieve the first cell from the first table (Word collections are one-based).
                Word.Table table = doc.Tables[1];
                Word.Cell wordCell = table.Cell(1, 1);
                CKCell ckCell = new CKCell(wordCell);

                // Verify that the CKCell wraps a valid COMCell.
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
                Word.Cell wordCell = table.Cell(1, 1);
                CKCell ckCell = new CKCell(wordCell);

                // Capture the original text.
                string originalText = ckCell.Text;
                string newText = originalText + " Test";

                // Update the text.
                ckCell.Text = newText;
                Assert.IsTrue(CKTextHelper.ScrunchEquals(newText, ckCell.Text), "The Text property should update to the new value.");

                // Optionally restore original text.
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
                Word.Cell wordCell = table.Cell(1, 1);
                CKCell ckCell = new CKCell(wordCell);

                // Save the original background color.
                Word.WdColor originalColor = ckCell.BackgroundColor;

                // Set a new background color (for example, red).
                ckCell.BackgroundColor = Word.WdColor.wdColorRed;
                Assert.AreEqual(Word.WdColor.wdColorRed, ckCell.BackgroundColor, "Background color should update to wdColorRed.");

                // Restore the original background color.
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
                Word.Cell wordCell = table.Cell(1, 1);
                CKCell ckCell = new CKCell(wordCell);

                // Save the original foreground color.
                Word.WdColor originalColor = ckCell.ForegroundColor;

                // Set a new foreground color (for example, blue).
                ckCell.ForegroundColor = Word.WdColor.wdColorBlue;
                Assert.AreEqual(Word.WdColor.wdColorBlue, ckCell.ForegroundColor, "Foreground color should update to wdColorBlue.");

                // Restore the original foreground color.
                ckCell.ForegroundColor = originalColor;
            });
        }
    }
}
