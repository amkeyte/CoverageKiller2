using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Tests for the ShadowWorkspace utility class.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0009
    /// </remarks>
    [TestClass]
    public class ShadowWorkspaceTests
    {
        //******* Standard Rigging ********
        public TestContext TestContext { get; set; }
        private string _testFilePath;
        private CKDocument _testFile;

        [TestInitialize]
        public void Setup()
        {
            Log.Information($"Running test => {GetType().Name}::{TestContext.TestName}");
            _testFilePath = RandomTestHarness.TestFile1;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }

        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: true);
            Log.Information($"Completed test => {GetType().Name}::{TestContext.TestName}; status: {TestContext.CurrentTestOutcome}");
        }
        //******* End Standard Rigging ********

        [TestMethod]
        public void GetShadowDocument_CreatesValidShadowWorkspace()
        {
            var app = _testFile.Application;
            using (var workspace = app.GetShadowWorkspace())
            {
                Assert.IsNotNull(workspace, "ShadowWorkspace instance was null.");
                Assert.IsNotNull(workspace.Document, "Internal CKDocument was null.");
                Assert.IsFalse(workspace.Document.IsOrphan, "Document was marked as orphan.");
            }
        }

        [TestMethod]
        public void CloneFrom_ToEnd_CreatesMatchingParagraph()
        {
            var para = _testFile.Sections[1].Paragraphs[1];
            var app = _testFile.Application;
            using (var shadow = app.GetShadowWorkspace())
            {
                var clone = shadow.CloneFrom(para);

                Assert.IsNotNull(clone);
                Assert.IsInstanceOfType(clone, typeof(CKParagraph));
                Assert.IsTrue(para.ScrunchedText == clone.ScrunchedText, "Cloned text did not match.");
            }
        }

        [TestMethod]
        public void CloneFrom_ToRange_CreatesMatchingParagraph()
        {
            var para = _testFile.Sections[1].Paragraphs[1];
            var app = _testFile.Application;
            using (var shadow = app.GetShadowWorkspace())
            {
                var range = shadow.Document.Range().CollapseToEnd();
                var clone = shadow.CloneFrom(para, range);

                Assert.IsNotNull(clone);
                Assert.IsInstanceOfType(clone, typeof(CKParagraph));
                Assert.IsTrue(para.ScrunchedText == clone.ScrunchedText, "Cloned text did not match.");
            }
        }

        [TestMethod]
        public void CloneFrom_ToCoordinates_CreatesMatchingParagraph()
        {
            var para = _testFile.Sections[1].Paragraphs[1];
            var app = _testFile.Application;
            using (var shadow = app.GetShadowWorkspace())
            {
                var start = shadow.Document.Range().End - 1;
                var end = start;
                var clone = shadow.CloneFrom(para, start, end);

                Assert.IsNotNull(clone);
                Assert.IsInstanceOfType(clone, typeof(CKParagraph));
                Assert.IsTrue(para.ScrunchedText == clone.ScrunchedText, "Cloned text did not match.");
            }
        }
        [TestMethod]
        public void ShadowWorkspace_Disposal_CleansUpDocument()
        {
            var app = _testFile.Application;
            CKDocument shadowDoc;
            string docLogId;

            using (var shadow = app.GetShadowWorkspace())
            {
                shadowDoc = shadow.Document;
                docLogId = shadowDoc.LogId;
                Assert.IsTrue(app.Documents.Contains(shadowDoc), "Shadow document should be tracked before disposal.");
            }

            // After disposal, document should no longer be present in the app.
            Assert.IsFalse(app.Documents.Any(d => d.LogId == docLogId), "Shadow document should be removed after disposal.");
        }

    }
}
