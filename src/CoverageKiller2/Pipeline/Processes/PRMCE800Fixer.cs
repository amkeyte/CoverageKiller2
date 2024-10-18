using CoverageKiller2.Pipeline.WordHelpers;
using Serilog;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class PRMCE800Fixer : CKWordPipelineProcess
    {
        private class _SS
        {
            public static readonly string BuildingNameF = "Building: _PRMCE 800";
            public static readonly string BuildingNameR = "Building: PRCME All Wings";
            public static readonly string FloorPlanF = "A?.*";
            public static readonly string FloorPlanR = "Wing {0}; Floor {1}";
            public static readonly string ChannelF = "852937";
            public static readonly string ChannelR = "852.93750";
            public static readonly string FloorSectionHeadingTable_F = "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)";
            public static readonly string FloorSectionHeadingTable_RemoveCols = "Ant Gain\tCable Loss\tPh.\tType\tMod\tNAC";
            public static readonly string FloorSectionHeadingTable_Band_F = "BAND";
            public static readonly int FloorSectionHeadingTable_Band_Row = 2;
            public static readonly string FloorSectionHeadingTable_Band_CellR = "800 Mhz";
            public static readonly string FloorSectionGridNotesTable_F = "Grid\t# of Areas\tArea Size (sq. ft)\tArea Width\r\n(ft)\tArea Height\r\n(ft)\tIgnore Area Color\tComments\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_F = "Critical Point\tDL\r\nPower\r\n(dBm)\tDL\r\nDAQ\tUL\r\nPower\r\n(dBm)\tUL\r\nDAQ\tResult\tDL\r\nLoss\r\n(dB)\tComment\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_ULPower = "DL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string FloorSectionAreaReportTable_F = "Grid\tArea\tDL\r\nPower\r\n(dBm)\tDL\r\nDAQ\tUL\r\nPower\r\n(dBm)\tUL\r\nDAQ\r\n\tResult\tDL\r\nLoss\r\n(dB)\tComment\r\n";
            public static readonly string FloorSectionAreaReportTable_ULPower = "DL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionAreaReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string SectionAdditionalInfo_F = "Additional Info";
        }

        public PRMCE800Fixer()
        {
        }

        public override void Process()
        {
            Log.Information("Fixing for PRMCE");

            Log.Debug("*** Building identifier replace");
            var textFinder = new TextFinder(CKDoc, _SS.BuildingNameF);

            while (textFinder.TryFind(out _, true))
            {
                textFinder.Replace(_SS.BuildingNameR);
            }

            Log.Debug("*** Floor names replace with rebuild.");
            textFinder = new TextFinder(CKDoc, _SS.FloorPlanF);

            while (textFinder.TryFind(out var foundRange2, true))
            {
                ReplaceFloorName(textFinder, foundRange2);
            }


            Log.Debug("*** fix floor section heading table.");
            foreach (var table in CKDoc.CKTables
                .Where(t => t.RowMatches(1, _SS.FloorSectionHeadingTable_F)))
                FixFloorSectionHeadingTable(table);

            Log.Debug("*** remove grid notes table.");
            foreach (var table in CKDoc.CKTables
                .Where(t => t.RowMatches(1, _SS.FloorSectionGridNotesTable_F))
                .Reverse())
                table.Delete();


            Log.Debug("*** remove extra critical point fields: ULPower, DL Loss");
            foreach (var table in CKDoc.CKTables
                .Where(t => t.RowMatches(2, _SS.FloorSectionCriticalPointReportTable_F)))
                FixFloorSectionCriticalPointReportTable(table);


            Log.Debug("*** remove extra area fields: ULPower, DL Loss");
            foreach (var table in CKDoc.CKTables
                .Where(t => t.RowMatches(2, _SS.FloorSectionAreaReportTable_F)))
                FixFloorSectionAreaReportTable(table);


            Log.Debug("*** remove end Info section");
            var infoSection = CKDoc.WordDoc.Sections.Last; //maybe someday create a CKSection
            textFinder = new TextFinder(CKDoc, _SS.FloorPlanF, infoSection.Range);

            // Check if we can find the text in the "Info" section
            if (textFinder.TryFind(out var foundText))
            {
                // If found, delete the section
                CKDoc.DeleteSection(infoSection.Index);
            }
        }

        private void FixFloorSectionCriticalPointReportTable(CKTable fixer)
        {
            Log.Debug("** Fixing table: {_SSID}", nameof(_SS.FloorSectionCriticalPointReportTable_F));
            fixer.RemoveColumnsByRowText(_SS.FloorSectionCriticalPointReportTable_ULPower, 2);
            fixer.RemoveColumnsByRowText(_SS.FloorSectionCriticalPointReportTable_DLLoss, 2);
            fixer.MakeFullPage();
        }
        private void FixFloorSectionAreaReportTable(CKTable fixer)
        {
            fixer.RemoveColumnsByRowText(_SS.FloorSectionAreaReportTable_ULPower, 2);
            fixer.RemoveColumnsByRowText(_SS.FloorSectionAreaReportTable_DLLoss, 2);
            fixer.MakeFullPage();
        }
        private static void FixFloorSectionHeadingTable(CKTable fixer)
        {
            var fixer = new CKWordTable(foundTable1);
            fixer.RemoveColumnsByHeader(_SS.FloorSectionHeadingTable_RemoveCols);
            fixer.SetCell(
                _SS.FloorSectionHeadingTable_Band_F,
                _SS.FloorSectionHeadingTable_Band_Row,
                _SS.FloorSectionHeadingTable_Band_CellR);
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
    }
}