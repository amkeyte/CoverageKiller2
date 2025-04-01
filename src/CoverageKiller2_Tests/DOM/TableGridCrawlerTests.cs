using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class TableGridCrawlerTests
    {

        public static List<string> DescribeTableRawCells(Word.Table table)
        {
            var lines = new List<string>();

            foreach (Word.Cell cell in table.Range.Cells)
            {
                string text = CKTextHelper.Scrunch(cell.Range.Text);
                lines.Add($"{text} ({cell.RowIndex},{cell.ColumnIndex})");
            }

            return lines;
        }


        private int testTable = 2;
        [TestMethod]
        public void Crawler_CrawlRows_ReturnsExpectedRows()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = doc.Tables[testTable];
                var crawler = new TableGridCrawler(table.COMTable);

                var rows = crawler.CrawlRows();
                foreach (var msg in DescribeTableRawCells(table.COMTable))
                    Debug.WriteLine($"{msg}");

                Debug.WriteLine("CrawlRows Result:");
                foreach (var row in rows)
                {
                    string message = string.Join(" | ", row.Select(c =>
                    {
                        string text = CKTextHelper.Scrunch(c.COMCell.Range.Text);
                        return $"{text}({c.GridRow},{c.GridCol})";
                    }));

                    Debug.Print(message);
                }

                Assert.IsTrue(rows.Count > 0);
                Assert.IsTrue(rows.All(r => r.Count > 0));
            });
        }

        [TestMethod]
        public void Crawler_CrawlColumns_ReturnsExpectedColumns()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = doc.Tables[testTable];
                var crawler = new TableGridCrawler(table.COMTable);

                var columns = crawler.CrawlColumns();

                foreach (var msg in DescribeTableRawCells(table.COMTable))
                    Debug.WriteLine($"{msg}");

                Debug.WriteLine("\r\rCrawlRows Result:");
                foreach (var col in columns)
                {
                    string message = string.Join(" | ", col.Select(c =>
                    {
                        string text = CKTextHelper.Scrunch(c.COMCell.Range.Text);
                        return $"{text}({c.GridRow},{c.GridCol})";
                    }));

                    Debug.Print(message);
                }

                Assert.IsTrue(columns.Count > 0);
                Assert.IsTrue(columns.All(c => c.Count > 0));
            });
        }
    }
}
