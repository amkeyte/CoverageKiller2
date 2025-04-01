using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKParagraphTests
    {
        [TestMethod]
        public void CKParagraph_Constructor_ThrowsArgumentNullException()
        {
            // Test that passing a null Word.Paragraph throws an ArgumentNullException.
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                var para = new CKParagraph(null);
            });
        }

        [TestMethod]
        public void CKParagraph_Constructor_WrapsParagraphSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // Ensure the document contains at least one paragraph.
                Word.Paragraph firstParagraph = doc.Paragraphs[1]; // Word collections are 1-based.
                CKParagraph ckParagraph = new CKParagraph(firstParagraph);

                // Validate that the COMParagraph property is set.
                Assert.IsNotNull(ckParagraph.COMParagraph, "COMParagraph should not be null.");

                // Validate that the inherited CKRange properties reflect the underlying paragraph's range.
                Assert.AreEqual(firstParagraph.Range.Start, ckParagraph.Start, "Start should match the underlying paragraph's range start.");
                Assert.AreEqual(firstParagraph.Range.End, ckParagraph.End, "End should match the underlying paragraph's range end.");
                Assert.AreEqual(firstParagraph.Range.Text, ckParagraph.Text, "Text should match the underlying paragraph's range text.");
            });
        }

        [TestMethod]
        public void CKParagraph_ToString_ReturnsValidString()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Word.Paragraph firstParagraph = doc.Paragraphs[1];
                CKParagraph ckParagraph = new CKParagraph(firstParagraph);

                string output = ckParagraph.ToString();
                // Check that the output includes "CKParagraph:" and contains the start and end positions.
                Assert.IsTrue(output.Contains("CKParagraph:"), "ToString() should include 'CKParagraph:'.");
                Assert.IsTrue(output.Contains(ckParagraph.Start.ToString()), "ToString() should include the Start value.");
                Assert.IsTrue(output.Contains(ckParagraph.End.ToString()), "ToString() should include the End value.");
            });
        }
    }
}
