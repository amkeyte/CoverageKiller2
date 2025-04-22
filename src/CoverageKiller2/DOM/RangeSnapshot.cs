using CoverageKiller2.Logging;
using K4os.Hash.xxHash;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
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
            this.Ping();
            TextPreview = CKTextHelper.Scrunch(range.Text);
            FontName = range.Font?.Name;
            FontSize = range.Font?.Size ?? 0;
            Alignment = range.ParagraphFormat?.Alignment ?? Word.WdParagraphAlignment.wdAlignParagraphLeft;
            object info = range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
            //PageNumber = info is int page ? page : -1;
            TableCount = range.Tables?.Count ?? 0;
            FieldCount = range.Fields?.Count ?? 0;

            FastHash = ComputeHash();
            this.Pong(msg: $"{nameof(FastHash)} = {FastHash}");
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
            LH.Ping<RangeSnapshot>();
            var result = a != null && b != null &&
                new RangeSnapshot(a).FastHash == new RangeSnapshot(b).FastHash;
            LH.Pong<RangeSnapshot>(msg: result.ToString());
            return result;
        }
    }
}

