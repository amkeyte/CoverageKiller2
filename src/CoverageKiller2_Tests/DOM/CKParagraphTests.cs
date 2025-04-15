using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Unit tests for <see cref="CKParagraph"/> wrapper.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0001
    /// </remarks>
    [TestClass]
    public class CKParagraphTests
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
        public void CKParagraph_Constructor_ThrowsArgumentNullException()
        {
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                var para = new CKParagraph(null, null);
            });
        }

        [TestMethod]
        public void CKParagraph_Constructor_WrapsParagraphSuccessfully()
        {
            var ckRange = _testFile.Sections[1];
            var ckParagraphs = ckRange.Paragraphs;
            Assert.IsTrue(ckParagraphs.Count > 0, "Test document should contain at least one paragraph.");

            var ckParagraph = ckParagraphs[1];

            Assert.IsNotNull(ckParagraph, "CKParagraph should not be null.");
            Assert.IsTrue(ckParagraph.Start >= 0, "Start should be a non-negative number.");
            Assert.IsTrue(ckParagraph.End >= ckParagraph.Start, "End should be greater than or equal to Start.");
            Assert.IsNotNull(ckParagraph.Text, "Text should not be null.");
        }

        [TestMethod]
        public void CKParagraph_ToString_ReturnsValidString()
        {
            var ckParagraph = _testFile.Sections[1].Paragraphs[1];
            string output = ckParagraph.ToString();

            Assert.IsTrue(output.Contains("CKParagraph:"), "ToString() should include 'CKParagraph:'.");
            Assert.IsTrue(output.Contains(ckParagraph.Start.ToString()), "ToString() should include the Start value.");
            Assert.IsTrue(output.Contains(ckParagraph.End.ToString()), "ToString() should include the End value.");
        }
    }
}
