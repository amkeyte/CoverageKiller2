using CoverageKiller2.DOM.Tables;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class MyWordExplorerClass
    {
        [TestMethod]
        public void SeeWhatWordDoes()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = doc.Tables[2];
                var sb = new StringBuilder();
                var cells = table.COMRange.Cells.ToList();

                foreach (var cell in cells)
                {
                    var cellText = CKTextHelper.Scrunch(cell.Range.Text);
                    int wdIndex = cells.IndexOf(cell);
                    sb.AppendLine($"Cell: {wdIndex} | coord:{cell.RowIndex},{cell.ColumnIndex} | '{cellText}' ");

                    string upResult = TryDirection(table, cells, cell, -1, 0, "Up");
                    string downResult = TryDirection(table, cells, cell, 2, 0, "Down");
                    string leftResult = TryDirection(table, cells, cell, 0, -1, "Left");
                    string rightResult = TryDirection(table, cells, cell, 0, 1, "Right");

                    sb.AppendLine(upResult);
                    sb.AppendLine(downResult);
                    sb.AppendLine(leftResult);
                    sb.AppendLine(rightResult);
                }

                Debug.Print(sb.ToString());
            });
        }

        private string TryDirection(CKTable table, System.Collections.Generic.List<Microsoft.Office.Interop.Word.Cell> cells, Microsoft.Office.Interop.Word.Cell origin, int rowOffset, int colOffset, string label)
        {
            try
            {
                _ = table.COMTable.Cell(origin.RowIndex, origin.ColumnIndex);
                var targetCell = table.COMTable.Cell(origin.RowIndex + rowOffset, origin.ColumnIndex + colOffset);
                targetCell = cells.FirstOrDefault(c => c.Range.COMEquals(targetCell.Range));

                if (targetCell == null || targetCell.Range.COMEquals(origin.Range))
                    throw new System.Runtime.InteropServices.COMException($"At {label} edge of table.");

                var text = CKTextHelper.Scrunch(targetCell.Range.Text);
                int index = cells.IndexOf(targetCell);
                return $"\t {label}: {index} | coord:{targetCell.RowIndex},{targetCell.ColumnIndex} | '{text}' ";
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return $"\t {label}: Exception: {ex.Message}";
            }
        }
    }
}
