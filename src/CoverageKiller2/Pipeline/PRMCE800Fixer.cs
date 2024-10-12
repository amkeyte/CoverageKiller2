//using Microsoft.Office.Interop.Word;

using Microsoft.Office.Interop.Word;
using Serilog;
using System;

namespace CoverageKiller2
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


        }

        public PRMCE800Fixer()
        {
        }

        public override void Process()
        {
            Log.Information("Fixing for PRMCE");

            //Building identifier replace
            var textFinder = new TextFinder(CKDoc, _SS.BuildingNameF);
            if (textFinder.TryFind(out _))
            {
                textFinder.Replace(_SS.BuildingNameR);

                while (textFinder.TryFindNext(out _, true))
                {
                    textFinder.Replace(_SS.BuildingNameR);

                }
            }

            //Floor names replace with rebuild.
            textFinder = new TextFinder(CKDoc, _SS.FloorPlanF);
            if (textFinder.TryFind(out var foundRange1))
            {
                ReplaceFloorName(textFinder, foundRange1);
                while (textFinder.TryFindNext(out var foundRange2, true))
                {
                    ReplaceFloorName(textFinder, foundRange2);
                }
            }

            //fix floor section heading table.
            var tableFinder = new TableFinder(CKDoc, _SS.FloorSectionHeadingTable_F);
            if (tableFinder.TryFind(out var foundTable1))
            {
                FixFloorSectionHeadingTable(foundTable1);

                while (tableFinder.TryFindNext(out var foundTable2))
                {
                    FixFloorSectionHeadingTable(foundTable2);
                }
            }
        }

        private static void FixFloorSectionHeadingTable(Table foundTable1)
        {
            var fixer = new TableFixer(foundTable1);
            fixer.RemoveColumnsByHeader(_SS.FloorSectionHeadingTable_RemoveCols);
            fixer.SetCell(
                _SS.FloorSectionHeadingTable_Band_F,
                _SS.FloorSectionHeadingTable_Band_Row,
                _SS.FloorSectionHeadingTable_Band_CellR);
        }

        private static void ReplaceFloorName(TextFinder tf, Range foundRange)
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

        public static (string, string) ExtractParts(string input)
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