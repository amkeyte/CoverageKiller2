using K4os.Hash.xxHash;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public static class WordInteropExtensions
    {
        public static IEnumerable<Word.Cell> AsEnumerable(this Word.Cells cells)
        {
            foreach (Word.Cell cell in cells)
            {
                yield return cell;
            }
        }
        public static List<Word.Cell> ToList(this Word.Cells cells)
        {
            return cells.AsEnumerable().ToList();

        }
        public static bool Contains(this Word.Range outer, Word.Range inner)
        {
            return inner.Start >= outer.Start && inner.End <= outer.End;
        }
        public static bool COMEquals(this Word.Range me, Word.Range other)
        {
            return me.Start >= other.Start && me.End <= other.End;
        }
    }
    public class RangeSnapshot
    {
        public string TextPreview { get; }
        public string FontName { get; }
        public float FontSize { get; }
        public Word.WdParagraphAlignment Alignment { get; }
        public int PageNumber { get; }
        public int TableCount { get; }
        public int FieldCount { get; }
        public ulong FastHash { get; }

        public RangeSnapshot(Word.Range range)
        {
            TextPreview = CKTextHelper.Scrunch(range.Text);
            FontName = range.Font?.Name;
            FontSize = range.Font?.Size ?? 0;
            Alignment = range.ParagraphFormat?.Alignment ?? Word.WdParagraphAlignment.wdAlignParagraphLeft;
            object info = range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
            //PageNumber = info is int page ? page : -1;
            TableCount = range.Tables?.Count ?? 0;
            FieldCount = range.Fields?.Count ?? 0;

            FastHash = ComputeHash();
        }

        private string TrimText(string text)
        {
            var trimmed = (text ?? "").Trim();
            return trimmed.Length > 100 ? trimmed.Substring(0, 100) : trimmed;
        }

        private ulong ComputeHash()
        {
            var data = $"{TextPreview}|{FontName}|{FontSize}|{Alignment}|{TableCount}|{FieldCount}";
            var bytes = Encoding.UTF8.GetBytes(data);
            return XXH64.DigestOf(bytes); // 64-bit fast, repeatable hash
        }

        public bool FastMatch(RangeSnapshot other)
        {
            return other != null && this.FastHash == other.FastHash;
        }

        public override string ToString()
        {
            return $"Snapshot(Hash={FastHash}, Text=\"{TextPreview}\")";
        }
        public bool FastMatch(Word.Range other)
        {
            return other != null &&
                FastHash == new RangeSnapshot(other).FastHash;
        }
        public static bool FastMatch(Word.Range a, Word.Range b)
        {
            return a != null && b != null &&
                new RangeSnapshot(a).FastHash == new RangeSnapshot(b).FastHash;
        }
    }
}

