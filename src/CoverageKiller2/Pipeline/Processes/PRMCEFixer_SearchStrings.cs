

using System.Collections.Generic;

namespace CoverageKiller2.Pipeline.Processes
{
    internal partial class PRMCEFixer : CKWordPipelineProcess
    {
        private _SS searchStrings = new _SS();
        private class _SS
        {
            public static readonly string BuildingNameR = "Building: PRCME All Wings";
            public static readonly string TestLocationR = "Colby Campus; 1700 13th Street, Everett, WA 98201";

            public static readonly string TestReportSummaryTableF = "Channel/ Ch Group\tFreq (MHz)\tTechnology\tBand\tResult\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n";
            public static readonly string TestReportSummaryChGrpF = "Channel/ Ch Group";
            public static readonly string TestReportSummaryBandColF = "Band";
            public static readonly string TestReportSummaryTechColF = "Technology";

            public static readonly string FloorPlanF = "A?.*";
            public static readonly string FloorPlanRoofF = "A.4R";

            public static readonly string FloorPlanR = "Wing {0}; Floor {1}";
            public static readonly string FloorSectionHeadingTable_F = "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)";
            public static readonly string FloorSectionHeadingTable_RemoveCols = "Ant Gain\tCable Loss\tPh.\tType\tMod\tNAC";
            public static readonly string FloorSectionHeadingTable_Band_F = "BAND";
            public static readonly int FloorSectionHeadingTable_Band_Row = 2;
            public static readonly string FloorSectionGridNotesTable_F = "Grid\t# of Areas\tArea Size (sq. ft)\tArea Width\r\n(ft)\tArea Height\r\n(ft)\tIgnore Area Color\tComments\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_F = "Critical Point Report";
            public static readonly string FloorSectionCriticalPointReportTable_ULPower = "UL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string FloorSectionAreaReportTable_F = "Area Report";
            public static readonly string FloorSectionAreaReportTable_ULPower = "UL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionAreaReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string SectionAdditionalInfo_F = "Additional Info";

            public List<ASubSS> Channels = new List<ASubSS>()
            {
                new SS_800(),
                new SS_450.CH1(),
                new SS_450.CH2(),
                new SS_450.CH3(),
                new SS_450.CH4(),
                new SS_450.CH5(),
                new SS_450.CH6(),
                new SS_450.CH7(),

            };

            public abstract class ASubSS
            {
                public abstract string TestReportSummaryTechColF { get; }
                public abstract string TestLocationF { get; }
                public abstract string BuildingNameF { get; }
                public abstract string ChannelF { get; }
                public abstract string ChannelR { get; }
                public abstract string FloorSectionHeadingTable_Band_CellR { get; }
                public abstract string TestReportSummaryBandColR { get; }
            }
            public class SS_800 : ASubSS
            {
                public override string TestLocationF => "_PRMCE 800";

                public override string BuildingNameF => "Building: _PRMCE 800";
                public override string ChannelF => "852937";
                public override string ChannelR => "852.93750";
                public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                public override string TestReportSummaryBandColR => "800 MHz Public Safety (SNO-911)";

                public override string TestReportSummaryTechColF => "P25";
            }
            public class SS_450
            {
                public class CH1 : ASubSS//451887
                {
                    public override string BuildingNameF => "Building: _PRMCE 451887";
                    public override string TestLocationF => "PRMCE 451887";
                    public override string ChannelF => "451887";
                    public override string ChannelR => "451.88750";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";

                    public override string TestReportSummaryTechColF => "MotoTRBO";
                }
                public class CH2 : ASubSS//452337
                {
                    public override string BuildingNameF => "Building: _PRMCE 452337";
                    public override string TestLocationF => "PRMCE 452337";
                    public override string ChannelF => "452337";
                    public override string ChannelR => "452.33750";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";
                    public override string TestReportSummaryTechColF => "MotoTRBO";

                }
                public class CH3 : ASubSS//452750
                {
                    public override string BuildingNameF => "Building: _PRMCE 452750";
                    public override string TestLocationF => "PRMCE 452750";
                    public override string ChannelF => "452750";
                    public override string ChannelR => "452.7500";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";
                    public override string TestReportSummaryTechColF => "MotoTRBO";

                }
                public class CH4 : ASubSS//461650
                {
                    public override string BuildingNameF => "Building: _PRMCE 461650";
                    public override string TestLocationF => "PRMCE 461650";
                    public override string ChannelF => "461650";
                    public override string ChannelR => "461.6500";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";
                    public override string TestReportSummaryTechColF => "MotoTRBO";

                }
                public class CH5 : ASubSS//461750
                {
                    public override string BuildingNameF => "Building: _PRMCE 461750";
                    public override string TestLocationF => "PRMCE 461750";
                    public override string ChannelF => "461750";
                    public override string ChannelR => "461.750";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";
                    public override string TestReportSummaryTechColF => "MotoTRBO";
                }
                public class CH6 : ASubSS//462112
                {
                    public override string BuildingNameF => "Building: _PRMCE 462112";
                    public override string TestLocationF => "PRMCE 462112";
                    public override string ChannelF => "462112";
                    public override string ChannelR => "462.1125";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";
                    public override string TestReportSummaryTechColF => "MotoTRBO";
                }
                public class CH7 : ASubSS//463550
                {
                    public override string BuildingNameF => "Building: _PRMCE 463550";
                    public override string TestLocationF => "PRMCE 463550";
                    public override string ChannelF => "463550";
                    public override string ChannelR => "463.5500";
                    public override string FloorSectionHeadingTable_Band_CellR => "450 MHz Commercial";
                    public override string TestReportSummaryBandColR => "450 MHz Commercial";

                    public override string TestReportSummaryTechColF => "MotoTRBO";

                }

            }
        }
    }
}
