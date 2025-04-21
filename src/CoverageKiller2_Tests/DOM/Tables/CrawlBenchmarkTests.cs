using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace CoverageKiller2.Tests.Benchmarking
{
    [TestClass]
    public class CrawlBenchmarkTests
    {
        //******* Standard Benchmark Rigging ********
        static int _testTableIndex = 16;
        static int _iterationCount = 1;

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
        public void Benchmark_CrawlVertically()
        {
            this.Ping("$$$");
            var benchmarkSw = Stopwatch.StartNew();

            var workspace = _testFile.Application.GetShadowWorkspace();
            workspace.ShowDebuggerWindow();

            var bootTime = benchmarkSw.ElapsedMilliseconds;

            var stageNames = new[]
            {
                "Total","CloneAndPrepare", "ParseText", "GetMasterGrid", "Normalize", "CrawlHoriz", "CrawlVertically"
            };
            var results = new Dictionary<string, List<long>>();
            foreach (var name in stageNames) results[name] = new List<long>();


            Log.Information("Benchmarking CrawlVertically workflow ({0} iterations) **************************************************", _iterationCount);

            for (int i = 0; i < _iterationCount; i++)
            {
                Log.Information("-- Iteration {0} of {1} -- ********************************", i + 1, _iterationCount);
                workspace.Document.Content.Delete();
                var ckTable = workspace.CloneFrom(_testFile.Tables[_testTableIndex]);
                var COMTable = ckTable.COMTable;
                var crawler = new GridCrawler5(COMTable);

                var swTot = Stopwatch.StartNew();
                var sw = Stopwatch.StartNew();
                var clonedTable = crawler.PrepareTable(COMTable);
                sw.Stop();
                results["CloneAndPrepare"].Add(sw.ElapsedMilliseconds);

                sw.Restart();
                var textGrid = crawler.ParseTableText(clonedTable);
                sw.Stop();
                results["ParseText"].Add(sw.ElapsedMilliseconds);

                sw.Restart();
                var masterGrid = crawler.GetMasterGrid(clonedTable);
                sw.Stop();
                results["GetMasterGrid"].Add(sw.ElapsedMilliseconds);

                sw.Restart();
                var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
                sw.Stop();
                results["Normalize"].Add(sw.ElapsedMilliseconds);

                sw.Restart();
                var horizontalGrid = crawler.CrawlHoriz(textGrid, normalizedGrid);
                sw.Stop();
                results["CrawlHoriz"].Add(sw.ElapsedMilliseconds);

                sw.Restart();
                var verticalGrid = crawler.CrawlVertically(0, textGrid, normalizedGrid);
                sw.Stop();
                results["CrawlVertically"].Add(sw.ElapsedMilliseconds);



                swTot.Stop();
                results["Total"].Add(swTot.ElapsedMilliseconds);
            }
            var bmTime = benchmarkSw.ElapsedMilliseconds;
            Log.Information("--- Benchmark Summary --- *************************************************************");
            Log.Information($"Boot up time: [{bootTime} ms]  |  Test run time: [{bmTime} ms]");
            foreach (var stage in stageNames)
            {
                var times = results[stage];
                var total = times.Sum();
                var avg = times.Average();
                Log.Information("{0,-18}: Total = {1,4} ms | Avg = {2,4:F1} ms", stage, total, avg);
            }
            this.Pong();
        }
        [TestMethod]
        public void Benchmark_AnalyzeTableRecursively()
        {
            this.Ping("$$$");
            var benchmarkSw = Stopwatch.StartNew();

            var workspace = _testFile.Application.GetShadowWorkspace();
            workspace.ShowDebuggerWindow();

            var bootTime = benchmarkSw.ElapsedMilliseconds;
            var results = new List<long>();
            Base1JaggedList<GridCell5> result = default;
            Log.Information("Benchmarking AnalyzeTableRecursively ({0} iterations) **************************************************", _iterationCount);

            for (int i = 0; i < _iterationCount; i++)
            {
                Log.Information("-- Iteration {0} of {1} -- ********************************", i + 1, _iterationCount);

                workspace.Document.Content.Delete();
                var ckTable = workspace.CloneFrom(_testFile.Tables[_testTableIndex]);
                var COMTable = ckTable.COMTable;
                var crawler = new GridCrawler5(COMTable);

                var sw = Stopwatch.StartNew();
                result = crawler.AnalyzeTableRecursively(COMTable);
                sw.Stop();

                results.Add(sw.ElapsedMilliseconds);
            }

            var bmTime = benchmarkSw.ElapsedMilliseconds;

            Log.Information("--- Dispatcher Benchmark Summary --- **************************************************");

            Log.Information(GridCrawler5.DumpGrid(result));


            Log.Information($"Boot up time: [{bootTime} ms]  |  Test run time: [{bmTime} ms]");

            long total = results.Sum();
            double avg = results.Average();
            Log.Information("AnalyzeTableRecursively: Total = {0} ms | Avg = {1:F1} ms", total, avg);

            this.Pong();
        }

    }
}
