using CoverageKiller2.DOM;
using CoverageKiller2.Tests;    // Contains LiveWordDocument
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.UnitTests
{
    [TestClass]
    public class CKRangeTests
    {

        [TestMethod]
        public void CKRange_Constructor_ThrowsArgumentNullException()
        {
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                // Passing null should throw.
                CKRange range = new CKRange(null);
            });
        }

        [TestMethod]
        public void CKRange_Equals_ReturnsTrueForSameUnderlyingRange()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Get a range (using the entire document content).
                Word.Range wordRange = doc.Content;
                CKRange ckRange1 = new CKRange(wordRange);
                CKRange ckRange2 = new CKRange(wordRange);

                // They wrap the same underlying COM range.
                Assert.AreEqual(ckRange1, ckRange2, "CKRange instances wrapping the same COMRange should be equal.");
                Assert.AreEqual(ckRange1.GetHashCode(), ckRange2.GetHashCode(), "Hash codes should be equal.");
            });
        }

        [TestMethod]
        public void CKRange_IsDirty_ReturnsFalseInitiallyAndTrueAfterModification()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Use a range that we can modify.
                Word.Range wordRange = doc.Content;
                CKRange ckRange = new CKRange(wordRange);

                // Initially, the range should not be dirty.
                Assert.IsFalse(ckRange.IsDirty, "New CKRange should not be dirty.");

                // Save the original text.
                string originalText = wordRange.Text;

                // Modify the underlying text.
                // Note: The document must be writable for this to work.
                wordRange.Text = originalText + " ";

                // Now, CKRange should be flagged as dirty.
                Assert.IsTrue(ckRange.IsDirty, "CKRange should be dirty after modifying the underlying text.");

                // Restore original text to keep the document intact.
                wordRange.Text = originalText;
            });
        }

        [TestMethod]
        public void CKRange_TextProperty_ReturnsUnderlyingText()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Word.Range wordRange = doc.Content;
                CKRange ckRange = new CKRange(wordRange);

                // The text property should mirror the COMRange's text.
                Assert.AreEqual(wordRange.Text, ckRange.Text, "CKRange.Text should match the underlying COMRange.Text.");
            });
        }

        [TestMethod]
        public void CKRange_WrapperProperties_ReturnNonNull()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Word.Range wordRange = doc.Content;
                CKRange ckRange = new CKRange(wordRange);

                // Check that the wrapper properties return non-null objects.
                Assert.IsNotNull(ckRange.Sections, "Sections property should not be null.");
                Assert.IsNotNull(ckRange.Paragraphs, "Paragraphs property should not be null.");
                Assert.IsNotNull(ckRange.Tables, "Tables property should not be null.");
                Assert.IsNotNull(ckRange.Cells, "Cells property should not be null.");
            });
        }
    }
}
