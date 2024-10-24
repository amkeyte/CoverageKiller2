using CoverageKiller2.Pipeline.WordHelpers;
using Serilog;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.Processes
{
    internal partial class PRMCEFixer : CKWordPipelineProcess
    {

        public PRMCEFixer()
        {
        }
        private _SS.ASubSS _ss = default;

        public override void Process()
        {
            Log.Information("Fixing for PRMCE");

            Log.Debug("*** Building identifier replace");


            string exTF = string.Empty;
            foreach (var ss in searchStrings.Channels)
            {
                TextFinder tf1 = new TextFinder(CKDoc, ss.BuildingNameF);
                exTF = tf1.SearchText;


                while (tf1.TryFind(out _, true))
                {
                    tf1.Replace(_SS.BuildingNameR);
                    //find active channel and use it downsteam
                    _ss = ss;
                }

                if (_ss != null) break;
            }

            if (_ss is null)
            {
                throw new Exception($"BuildingNameF not found: {exTF}");

            }
            Log.Debug("*** Floor names replace with rebuild.");
            var tf2 = new TextFinder(CKDoc, _SS.FloorPlanF);

            while (tf2.TryFind(out var foundRange2, true))
            {
                ReplaceFloorName(tf2, foundRange2);
            }



            Log.Debug("*** fix floor section heading table.");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(1, _SS.FloorSectionHeadingTable_F))
                .Reverse())
                FixFloorSectionHeadingTable(table);

            Log.Debug("*** remove grid notes table.");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(1, _SS.FloorSectionGridNotesTable_F))
                .Reverse())
                table.Delete();


            Log.Debug("*** remove extra critical point fields: ULPower, DL Loss");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(2, _SS.FloorSectionCriticalPointReportTable_F))
                .Reverse())
                FixFloorSectionCriticalPointReportTable(table);


            Log.Debug("*** remove extra area fields: ULPower, DL Loss");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(2, _SS.FloorSectionAreaReportTable_F))
                .Reverse())
                FixFloorSectionAreaReportTable(table);


            Log.Debug("*** remove end Info section");
            var infoSection = CKDoc.COMObject.Sections.Last; //maybe someday create a CKSection
            var tf3 = new TextFinder(CKDoc, _SS.FloorPlanF, infoSection.Range);

            // Check if we can find the text in the "Info" section
            if (tf3.TryFind(out var foundText))
            {
                // If found, delete the section
                CKDoc.DeleteSection(infoSection.Index);
            }
        }

        private void FixFloorSectionCriticalPointReportTable(CKTable fixer)
        {
            Log.Debug("** Fixing table: {_SSID}", nameof(_SS.FloorSectionCriticalPointReportTable_F));


            Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                nameof(PRMCEFixer), nameof(FixFloorSectionCriticalPointReportTable),
                $"{nameof(fixer)}({nameof(CKTable)}.{nameof(fixer.Index)}) --> ", fixer.Index));


            fixer.Rows.First().Delete();

            fixer.Columns.Reverse()
                 .First(col => col.Cells[1].Text == _SS.FloorSectionCriticalPointReportTable_ULPower)
                 .Delete();

            fixer.Columns.Reverse()
                 .First(col => col.Cells[1].Text == _SS.FloorSectionCriticalPointReportTable_DLLoss)
                 .Delete();

            fixer.AddAndMergeFirstRow("Critical Points");
            fixer.MakeFullPage();
        }
        private void FixFloorSectionAreaReportTable(CKTable fixer)
        {
            Log.Debug("** [BYPASSED] Fixing table: {_SSID}", nameof(_SS.FloorSectionAreaReportTable_F));
            //fixer.Columns
            //     .First(col => col.Cells[2].Text == _SS.FloorSectionAreaReportTable_ULPower)
            //     .Delete();

            //fixer.Columns
            //     .First(col => col.Cells[2].Text == _SS.FloorSectionAreaReportTable_DLLoss)
            //     .Delete();

            //fixer.MakeFullPage();
        }
        private void FixFloorSectionHeadingTable(CKTable fixer)
        {
            Log.Debug("TRACE => {func}({param1} = {pVal1})",
                nameof(FixFloorSectionHeadingTable),
                nameof(fixer),
                $"Table[{fixer.Index}]");

            Log.Debug("** Fixing table: {_SSID}", nameof(_SS.FloorSectionHeadingTable_F));

            var headersToRemove = _SS.FloorSectionHeadingTable_RemoveCols
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => NormalizeMatchString(s))
                .Reverse()
                .ToList();



            fixer.Columns
                .Where(col => headersToRemove
                    .Contains(NormalizeMatchString(col.Cells[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            fixer.SetCell(
                _SS.FloorSectionHeadingTable_Band_F,
                _SS.FloorSectionHeadingTable_Band_Row,
                _ss.FloorSectionHeadingTable_Band_CellR);

            fixer.MakeFullPage();
        }

        private static void ReplaceFloorName(TextFinder tf, Word.Range foundRange)
        {
            var x = ExtractParts(foundRange.Text);
            switch (x.Item1)
            {
                case "A1":
                    x.Item1 = "A, B, C";
                    break;
                case "A4":
                    x.Item1 = "D";
                    break;
                default:
                    throw new ArgumentException("A valid wing code was not found.");
            }

            string replaceText = string.Format(_SS.FloorPlanR, x.Item1, x.Item2);
            tf.Replace(replaceText);
        }

        private static (string, string) ExtractParts(string input)
        {
            // Find the index of the dot
            int dotIndex = input.IndexOf('.');

            if (dotIndex == -1)
            {
                throw new ArgumentException($"Input '{input}' does not contain a dot.");
            }

            // Extract the parts before and after the dot
            string part1 = input.Substring(0, dotIndex); // Part before the dot
            string part2 = input.Substring(dotIndex + 1); // Part after the dot

            return (part1, part2);
        }


        private static int _dbgCounter_NormalizeMatchString = 0;
        private static string NormalizeMatchString(string input)
        {
            //Log.Debug("Called ({_dbgCounter_NormalizeMatchString}): {nms}(input: {input})",
            //    _dbgCounter_NormalizeMatchString++,
            //    nameof(NormalizeMatchString),
            //    input);

            return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }
    }
}