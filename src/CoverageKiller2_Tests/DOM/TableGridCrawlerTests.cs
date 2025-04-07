using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.DOM.Tables
{

    [TestClass]
    public class TableGridCrawlerTests
    {
        [TestMethod("Don't forget me!")]
        public void DontForgetMe()
        {
            //stupid tag so I can not lose track of these commented tests.
            //tests are commented because of DOM rebuild.
        }
    }
}

//public static List<string> DescribeTableRawCells(Word.Table table)
//        {
//            var lines = new List<string>();

//            foreach (Word.Cell cell in table.Range.Cells)
//            {
//                string text = CKTextHelper.Scrunch(cell.Range.Text);
//                lines.Add($"{text} ({cell.RowIndex},{cell.ColumnIndex})");
//            }

//            return lines;
//        }
//        [TestMethod]
//        public void GetBottomRightCell_Returns_Correct_Cell()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[testTable];
//                var crawler = new TableGridCrawler(table.COMTable);
//                foreach (var msg in DescribeTableRawCells(table.COMTable))
//                    Debug.WriteLine($"{msg}");

//                var cell = crawler.GetBottomRightCell();
//                Debug.WriteLine($"\nBottom Right: {cell.Range.Text} ({cell.RowIndex},{cell.ColumnIndex})");
//            });
//        }
//        private int testTable = 2;
//        [TestMethod]
//        public void Crawler_CrawlRowsReverse_ReturnsExpectedRows()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[testTable];
//                var crawler = new TableGridCrawler(table.COMTable);

//                var rows = crawler.CrawlRowsReverse();
//                foreach (var msg in DescribeTableRawCells(table.COMTable))
//                    Debug.WriteLine($"{msg}");

//                Debug.WriteLine("CrawlRows Result:");
//                foreach (var row in rows)
//                {
//                    string message = string.Join(" | ", row.Select(c =>
//                    {
//                        string text = CKTextHelper.Scrunch(c.COMCell.Range.Text);
//                        return $"{text}({c.GridRow},{c.GridCol})";
//                    }));

//                    Debug.Print(message);
//                }

//                Assert.IsTrue(rows.Count > 0);
//                Assert.IsTrue(rows.All(r => r.Count > 0));
//            });
//        }
//        [TestMethod]
//        public void Crawler_CrawlRows_ReturnsExpectedRows()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[testTable];
//                var crawler = new TableGridCrawler(table.COMTable);

//                var rows = crawler.CrawlRows();
//                foreach (var msg in DescribeTableRawCells(table.COMTable))
//                    Debug.WriteLine($"{msg}");

//                Debug.WriteLine("CrawlRows Result:");
//                foreach (var row in rows)
//                {
//                    string message = string.Join(" | ", row.Select(c =>
//                    {
//                        string text = CKTextHelper.Scrunch(c.COMCell.Range.Text);
//                        return $"{text}({c.GridRow},{c.GridCol})";
//                    }));

//                    Debug.Print(message);
//                }

//                Assert.IsTrue(rows.Count > 0);
//                Assert.IsTrue(rows.All(r => r.Count > 0));
//            });
//        }

//        [TestMethod]
//        public void Crawler_CrawlColumns_ReturnsExpectedColumns()
//        {
//            LiveWordDocument.WithTestDocument(doc =>
//            {
//                var table = doc.Tables[testTable];
//                var crawler = new TableGridCrawler(table.COMTable);

//                var columns = crawler.CrawlColumns();

//                foreach (var msg in DescribeTableRawCells(table.COMTable))
//                    Debug.WriteLine($"{msg}");

//                Debug.WriteLine("\r\rCrawlRows Result:");
//                foreach (var col in columns)
//                {
//                    string message = string.Join(" | ", col.Select(c =>
//                    {
//                        string text = CKTextHelper.Scrunch(c.COMCell.Range.Text);
//                        return $"{text}({c.GridRow},{c.GridCol})";
//                    }));

//                    Debug.Print(message);
//                }

//                Assert.IsTrue(columns.Count > 0);
//                Assert.IsTrue(columns.All(c => c.Count > 0));
//            });
//        }
//    }
//}
