using CoverageKiller2.Logging;
using K4os.Hash.xxHash;
using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a cached, COM-free fingerprint of a Word.Range.
    /// Supports fast hash comparisons and optional slow fallbacks.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0001
    /// </remarks>
    public class RangeSnapshot
    {
        public string TextPreview { get; }
        public string FontName { get; }
        public float FontSize { get; }
        public Word.WdParagraphAlignment Alignment { get; }
        public int TableCount { get; }
        public int FieldCount { get; }
        public string FastHash { get; }

        public RangeSnapshot(Word.Range range)
        {
            LH.Debug("Tracker[!sd]");

            if (range == null) throw new ArgumentNullException(nameof(range));

            TextPreview = CKTextHelper.Scrunch(range.Text);
            FontName = range.Font?.Name;
            FontSize = range.Font?.Size ?? 0;
            Alignment = range.ParagraphFormat?.Alignment ?? Word.WdParagraphAlignment.wdAlignParagraphLeft;
            TableCount = range.Tables?.Count ?? 0;
            FieldCount = range.Fields?.Count ?? 0;

            var hash = ComputeHash().ToString();
            FastHash = hash.Substring(hash.Length - 6);
        }

        private ulong ComputeHash()
        {
            LH.Debug("Tracker[!sd]");
            var data = $"{TextPreview}|{FontName}|{FontSize}|{Alignment}|{TableCount}|{FieldCount}";
            var bytes = Encoding.UTF8.GetBytes(data);
            return XXH64.DigestOf(bytes);
        }

        /// <summary>
        /// Fast comparison using only precomputed hash values.
        /// </summary>
        public bool FastMatch(RangeSnapshot other)
        {
            LH.Debug("Tracker[!sd]");
            return other != null && this.FastHash == other.FastHash;
        }

        /// <summary>
        /// Builds a new snapshot from the given Word.Range and compares hashes.
        /// Accepts COM access.
        /// </summary>
        public bool SlowMatch(Word.Range other)
            => other != null && FastMatch(new RangeSnapshot(other));

        /// <summary>
        /// Uses the cached snapshot from a CKRange and falls back to SlowMatch if unavailable.
        /// </summary>
        public bool Match(CKRange range)
        {
            if (range == null) return false;

            var otherSnapshot = range.Snapshot;
            if (otherSnapshot != null && FastMatch(otherSnapshot))
                return true;

            return SlowMatch(range.COMRange);
        }
        /// <summary>
        /// Static form of FastMatch that compares two existing snapshots.
        /// No COM access.
        /// </summary>
        public static bool FastMatch(RangeSnapshot a, RangeSnapshot b)
            => a != null && b != null && a.FastMatch(b);

        /// <summary>
        /// Static form of SlowMatch that builds snapshots from Word.Range objects.
        /// Uses COM.
        /// </summary>
        public static bool SlowMatch(Word.Range a, Word.Range b)
        {
            LH.Debug("Tracker[!sd]");
            if (a == null || b == null) return false;
            var snapA = new RangeSnapshot(a);
            var snapB = new RangeSnapshot(b);
            return snapA.FastMatch(snapB);
        }

        /// <summary>
        /// Static form of Match that compares a snapshot against a CKRange, falling back if needed.
        /// </summary>
        public static bool Match(RangeSnapshot a, CKRange b)
            => a != null && b != null && a.Match(b);

        /// <summary>
        /// Static form of Match that compares two CKRanges, using snapshot fastmatch with fallback.
        /// </summary>
        public static bool Match(CKRange a, CKRange b)
        {
            if (a == null || b == null) return false;

            var snapA = a.Snapshot ?? new RangeSnapshot(a.COMRange);
            var snapB = b.Snapshot ?? new RangeSnapshot(b.COMRange);

            return snapA.FastMatch(snapB);
        }

        public override string ToString()
            => $"{FastHash}";
    }
}
