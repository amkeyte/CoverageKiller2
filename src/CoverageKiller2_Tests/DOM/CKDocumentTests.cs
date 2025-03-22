using CoverageKiller2.Tests;     // Contains VstoTestDocumentLoader.
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKDocumentTests
    {
        // Update the path below to the location of your test document.


        [TestMethod]
        public void CKDocument_Constructor_LoadsDocumentSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Create an instance of CKDocument from the live Word.Document.
                CKDocument ckDoc = new CKDocument(doc);

                // Assert that the document was loaded and processed correctly.
                Assert.IsNotNull(ckDoc, "CKDocument should not be null.");
                Assert.IsTrue(ckDoc.COMObject.Paragraphs.Count > 0, "The document should contain paragraphs.");
            });
        }

        [TestMethod]
        public void CKDocument_Properties_ReturnExpectedValues()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Create the CKDocument instance.
                CKDocument ckDoc = new CKDocument(doc);

                // Test that the COMObject property is not null.
                Assert.IsNotNull(ckDoc.COMObject, "COMObject should not be null.");

                // Test that the WordApp property returns the Application from the COMObject.
                Assert.IsNotNull(ckDoc.WordApp, "WordApp should not be null.");
                Assert.AreEqual(ckDoc.COMObject.Application, ckDoc.WordApp, "WordApp should match COMObject.Application.");

                // Test that the Content property is not null.
                Assert.IsNotNull(ckDoc.Content, "Content should not be null.");

                // Test that the FullPath property matches the provided document path.
                Assert.AreEqual(LiveWordDocument.Default, ckDoc.FullPath, "FullPath should match the test document path.");
            });
        }


    }
}
