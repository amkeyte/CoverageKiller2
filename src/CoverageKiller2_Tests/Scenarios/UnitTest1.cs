using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.Tests.Scenarios
{
    [TestClass]
    public class RandomTestHarnessTests
    {

        [TestClass]
        public class TempFilePreservationTests
        {
            public TestContext TestContext { get; set; }
            private CKDocument _testFile;

            [TestInitialize]
            public void Setup()
            {
                Log.Information($"Running test => {GetType().Name}::{TestContext.TestName}");

                RandomTestHarness.PreserveTempFilesAfterTest = true;

                _testFile = RandomTestHarness.GetTempDocumentFrom(RandomTestHarness.TestFile1);

                // Key: pin it alive so CKApplication won't close it later
                _testFile.KeepAlive = true;
            }

            [TestCleanup]
            public void Cleanup()
            {
                // Always reset the flag so next tests don't accidentally inherit
                RandomTestHarness.PreserveTempFilesAfterTest = false;

                RandomTestHarness.CleanUp(_testFile);
            }

            [TestMethod]
            public void MyManualInspectionTest()
            {
                // Your test logic here
                // At the end of the test, Word will remain open with your temp file available to inspect!
            }
        }
    }
}
