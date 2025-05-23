﻿using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKRangeTests
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
        public void CKRange_SetText_OnMixedRange_ThrowsCOMException()
        {
            var table = _testFile.Tables.FirstOrDefault();
            Assert.IsNotNull(table);

            int start = table.Start + 3;
            int end = table.End + 20;

            var mixedRange = _testFile.Range(start, end);

            Assert.ThrowsException<COMException>(() =>
            {
                mixedRange.Text = "Test new text";
            });
        }

        [TestMethod]
        public void CKRange_Cache_AutomaticallyResolvesText()
        {
            var range = _testFile.Range(30, 40);

            string original = range.Text;
            string newText = original + " extra";

            range.Text = newText;

            Assert.IsTrue(range.IsDirty, "After changing text, range should be dirty.");

            // Cache access (Text) should internally trigger refresh
            var text = range.Text;

            Assert.IsFalse(range.IsDirty, "Accessing text should have cleared dirty flag.");
            Assert.AreEqual(newText, text);
            Assert.AreEqual(CKTextHelper.Scrunch(newText), range.ScrunchedText);
        }

        [TestMethod]
        public void CollapseToEnd_ShouldReturnCollapsedRange()
        {
            var range = _testFile.Range(10, 30);
            var collapsed = range.CollapseToEnd();

            Assert.AreEqual(collapsed.Start, collapsed.End);
            Assert.AreEqual(collapsed.Start, range.End);
        }

        [TestMethod]
        public void CollapseToStart_ShouldReturnCollapsedRange()
        {
            var range = _testFile.Range(10, 30);
            var collapsed = range.CollapseToStart();

            Assert.AreEqual(collapsed.Start, collapsed.End);
            Assert.AreEqual(collapsed.Start, range.Start);
        }

        [TestMethod]
        public void FormattedText_Set_ShouldApplyFormatting()
        {
            var range = _testFile.Range(5, 15);
            var target = _testFile.Range(40, 50);

            target.FormattedText = range.FormattedText;

            // Auto-refresh through accessing text
            var text = target.Text;

            Assert.AreEqual(CKTextHelper.Scrunch(range.Text), CKTextHelper.Scrunch(text));
        }

        [TestMethod]
        public void CKRange_Cells_ShouldReturnCollection()
        {
            var table = _testFile.Tables[1];
            var range = new CKRange(table.COMTable.Range, _testFile);

            var cells = range.Cells;

            Assert.IsNotNull(cells);
            Assert.IsTrue(cells.Count > 0);
        }

        [TestMethod]
        public void CKRange_EqualityAndHashCode_ShouldBeConsistent()
        {
            var rangeA = _testFile.Range(10, 20);
            var rangeB = _testFile.Range(10, 20);
            var rangeC = _testFile.Range(20, 30);

            Assert.AreEqual(rangeA, rangeB);
            Assert.AreNotEqual(rangeA, rangeC);

            Assert.AreEqual(rangeA.GetHashCode(), rangeB.GetHashCode());
            Assert.AreNotEqual(rangeA.GetHashCode(), rangeC.GetHashCode());
        }

        [TestMethod]
        public void CKRange_DeferCOM_StartsDirty()
        {
            var deferredRange = new CKRange(_testFile); // Using defer constructor
            Assert.IsTrue(deferredRange.IsDirty, "Deferred range should be initially dirty.");
        }

        [TestMethod]
        public void CKRange_DeferCOM_CacheLiftsDefer()
        {
            var deferredRange = new CKRange(_testFile); // Defer mode
            Assert.IsTrue(deferredRange.IsDirty, "Deferred range should start dirty.");

            Assert.ThrowsException<InvalidOperationException>(() =>
            {
                var text = deferredRange.Text; // Should throw because no COMRange
            }, "Accessing Text on a truly empty deferred range should throw due to missing COMRange.");
        }

        [TestMethod]
        public void CKRange_DeferCOM_IsDirtyWithoutLifting()
        {
            var deferredRange = new CKRange(_testFile); // Defer mode

            bool dirty = deferredRange.IsDirty;
            Assert.IsTrue(dirty, "Deferred range should report dirty without lifting defer.");
        }
    }
}
