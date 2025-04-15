using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.Test
{
    [TestClass]
    public static class TestSetup
    {
        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            // This forces static constructor of RandomTestHarness to run
            _ = RandomTestHarness.Application;
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            RandomTestHarness.Shutdown();
        }
    }
}
