using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
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
        public Tracer Tracer { get; } = new Tracer(typeof(PRMCEFixer));
        public PRMCEFixer(string docType)
        {
            Tracer.Enabled = false;

            switch (docType)
            {
                case "800Mhz":
                    _ss = new _SS.SS_800();
                    break;
                case "UHF":
                    _ss = new _SS.SS_UHF();
                    break;
                default:
                    break;
            }
        }
        private _SS.ASubSS _ss = default;

        public override void Process()
        {
            Log.Information("**** Fixing for PRMCE");

            Log.Information("*** Various text relacements");


            ReplaceBuildingID();
            ReplaceTestLocaton();
            ReplaceChannel();

            Log.Information("*** Editing Test Report Summary");

            var TRSTable = CKDoc.Tables
                .First(t => t.RowMatches(1, _ss.TestReportSummaryTableF));

            TRSTable.Columns[1].Delete();

            TRSTable.MakeFullPage();

            TRSTable.Columns
                .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.TestReportSummaryBandColF))
                .Cells[2].Text = _ss.TestReportSummaryBandColR;
            TRSTable.Columns
                .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.TestReportSummaryTechColF))
                .Cells[2].Text = _ss.TestReportSummaryTechColR;

            Log.Information("*** Fix Threshold settings table.");

            var ThresholdSettingsTable = CKDoc.Tables
                .First(t => t.RowMatches(1, _ss.ThresholdSettingsTable_F));

            ThresholdSettingsTable.SetCell(
                _ss.ThresholdSettingsTable_Measurement_F,
                2, // hack
                _ss.ThresholdSettingsTable_Measurement_R);


            Log.Information("*** Floor names replace with rebuild.");
            var tf2 = new TextFinder(CKDoc, _ss.FloorPlanF);

            while (tf2.TryFind(out var foundRange2, true))
            {
                ReplaceFloorName(tf2, foundRange2);
            }

            var tf2A = new TextFinder(CKDoc, "A.4R");
            while (tf2A.TryFind(out var foundRange2A))
            {
                foundRange2A.Text = "Wing D; Roof";
            }


            Log.Information("*** fix floor section heading table.");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(1, _ss.FloorSectionHeadingTable_F))
                .Reverse())
                FixFloorSectionHeadingTable(table);

            Log.Information("*** remove grid notes table.");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(1, _ss.FloorSectionGridNotesTable_F))
                .Reverse())
            {
                table.Tracer.Stash(nameof(table.Index), table.Index);
                table.Delete();
            }


            Log.Information("*** remove extra critical point fields: ULPower, DL Loss");
            foreach (var table in CKDoc.Tables
                .Where(t => t.RowMatches(1, _ss.FloorSectionCriticalPointReportTable_F)))
            {

                FixFloorSectionCriticalPointReportTable(table);
            }


            Log.Information("*** remove extra area fields: ULPower, DL Loss");
            foreach (var table in CKDoc.Tables
                .Where(t => NormalizeMatchStrings(
                    t.Rows[1].Cells
                        .Aggregate("", (acc, c) => acc + c.Text), _ss.FloorSectionAreaReportTable_F)))
            {
                FixFloorSectionAreaReportTable(table);
            }



            Log.Information("*** remove end Info section");

            Tracer.Log("DataPoints", new DataPoints()
                .Add("CKDoc.COMObject.Sections.Count", CKDoc.COMObject.Sections.Count));

            var infoSection = CKDoc.COMObject.Sections.Last; //maybe someday create a CKSection
            var tf3 = new TextFinder(CKDoc, _ss.SectionAdditionalInfo_F, infoSection.Range);

            // Check if we can find the text in the "Info" section, because it's possibly already removed.
            if (tf3.TryFind(out var foundText))
            {
                // If found, delete the section
                CKDoc.DeleteSection(infoSection.Index);
            }
            else
            {
                Tracer.Log("Section was not deleted.");
            }
        }

        private void ReplaceChannel()
        {
            //HACK: do location firt to get _ss
            TextFinder tf1 = new TextFinder(CKDoc, _ss.ChannelF);

            while (tf1.TryFind(out _, true))
            {
                tf1.Replace(_ss.ChannelR);

            }
        }

        private void ReplaceTestLocaton()
        {
            //HACK: do location firt to get _ss
            TextFinder tf1 = new TextFinder(CKDoc, _ss.TestLocationF);

            while (tf1.TryFind(out _, true))
            {
                tf1.Replace(_ss.TestLocationR);

            }


        }
        private void ReplaceBuildingID()
        {
            //foreach (var ss in searchStrings.Channels)
            //{
            TextFinder tf1 = new TextFinder(CKDoc, _ss.BuildingNameF);

            while (tf1.TryFind(out _, true))
            {
                tf1.Replace(_ss.BuildingNameR);
                //_ss = ss;
            }

            //    if (_ss != null) break;
            //}

            //if (_ss is null)
            //{
            //    throw new Exception($"BuildingNameF not found.");
            //}
        }
        private void FixFloorSectionCriticalPointReportTable(CKTable fixer)
        {

            Tracer.Log("Entering", "**", new DataPoints()
                .Add($"{nameof(fixer)}.Index", fixer.Index));

            try
            {
                Tracer.Log("Deleting first row");

                fixer.Rows.First().Delete();
            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }

            try
            {
                Tracer.Log("Deleting columns UL Power");

                fixer.Columns
                     .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.FloorSectionCriticalPointReportTable_ULPower))
                     .Delete();

            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }


            try
            {
                Tracer.Log("Deleting columns B");

                fixer.Columns
                     .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.FloorSectionCriticalPointReportTable_DLLoss))
                     .Delete();
            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }
            try
            {
                Tracer.Log("Returing first row and making fix width");

                fixer.AddAndMergeFirstRow("Critical Point Report");
                fixer.MakeFullPage();
            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }
        }
        private void FixFloorSectionAreaReportTable(CKTable fixer)
        {
            Tracer.Log("Entering", "**", new DataPoints($"{nameof(fixer)}.Index", fixer.Index));


            Tracer.Log("Deleting first row");
            fixer.Rows.First().Delete();

            fixer.Columns
                 .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.FloorSectionAreaReportTable_ULPower))
                 .Delete();

            fixer.Columns
                 .First(col => NormalizeMatchStrings(col.Cells[1].Text, _ss.FloorSectionAreaReportTable_DLLoss))
                 .Delete();

            fixer.AddAndMergeFirstRow("Grid Area Report");

            fixer.MakeFullPage();
        }
        private void FixFloorSectionHeadingTable(CKTable fixer)
        {
            var headersToRemove = _ss.FloorSectionHeadingTable_RemoveCols
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => NormalizeMatchString(s))
                .Reverse()
                .ToList();



            fixer.Columns
                .Where(col => headersToRemove
                    .Contains(NormalizeMatchString(col.Cells[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            fixer.SetCell(
                _ss.FloorSectionHeadingTable_Band_F,
                _ss.FloorSectionHeadingTable_Data_Row,
                _ss.FloorSectionHeadingTable_Band_CellR);

            fixer.SetCell(
                _ss.FloorSectionHeadingTable_Tech_F,
                _ss.FloorSectionHeadingTable_Data_Row,
                _ss.FloorSectionHeadingTable_Tech_CellR);

            fixer.MakeFullPage();
        }






        private void ReplaceFloorName(TextFinder tf, Word.Range foundRange)
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

            string replaceText = string.Format(_ss.FloorPlanR, x.Item1, x.Item2);
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
            return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }
        private bool NormalizeMatchStrings(string str1, string str2)
        {
            return NormalizeMatchString(str1) == NormalizeMatchString(str2);
        }
    }
}