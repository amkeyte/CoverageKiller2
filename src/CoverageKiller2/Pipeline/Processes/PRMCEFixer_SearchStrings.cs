

using System.Collections.Generic;

namespace CoverageKiller2.Pipeline.Processes
{
    internal partial class PRMCEFixer : CKWordPipelineProcess
    {
        private _SS searchStrings = new _SS();
        private class _SS
        {
            public static readonly string BuildingNameR = "Building: PRCME All Wings";
            public static readonly string FloorPlanF = "A?.*";
            public static readonly string FloorPlanR = "Wing {0}; Floor {1}";
            public static readonly string FloorSectionHeadingTable_F = "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)";
            public static readonly string FloorSectionHeadingTable_RemoveCols = "Ant Gain\tCable Loss\tPh.\tType\tMod\tNAC";
            public static readonly string FloorSectionHeadingTable_Band_F = "BAND";
            public static readonly int FloorSectionHeadingTable_Band_Row = 2;
            public static readonly string FloorSectionGridNotesTable_F = "Grid\t# of Areas\tArea Size (sq. ft)\tArea Width\r\n(ft)\tArea Height\r\n(ft)\tIgnore Area Color\tComments\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_F = "Critical Point\tDL\r\nPower\r\n(dBm)\tDL\r\nDAQ\tUL\r\nPower\r\n(dBm)\tUL\r\nDAQ\tResult\tDL\r\nLoss\r\n(dB)\tComment\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_ULPower = "DL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionCriticalPointReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string FloorSectionAreaReportTable_F = "Grid\tArea\tDL\r\nPower\r\n(dBm)\tDL\r\nDAQ\tUL\r\nPower\r\n(dBm)\tUL\r\nDAQ\r\n\tResult\tDL\r\nLoss\r\n(dB)\tComment\r\n";
            public static readonly string FloorSectionAreaReportTable_ULPower = "DL\r\nPower\r\n(dBm)\r\n";
            public static readonly string FloorSectionAreaReportTable_DLLoss = "DL\r\nLoss\r\n(dB)\r\n";
            public static readonly string SectionAdditionalInfo_F = "Additional Info";

            public List<ASubSS> Channels = new List<ASubSS>()
            {
                new SS_800(),
                //new SS_450.CH1(),
                //new SS_450.CH2(),
                //new SS_450.CH3(),
                //new SS_450.CH4(),
                //new SS_450.CH5(),
                //new SS_450.CH6(),
                //new SS_450.CH7(),

            };

            public abstract class ASubSS
            {
                public abstract string BuildingNameF { get; }
                public abstract string ChannelF { get; }
                public abstract string ChannelR { get; }
                public abstract string FloorSectionHeadingTable_Band_CellR { get; }
            }
            public class SS_800 : ASubSS
            {
                public override string BuildingNameF => "Building: _PRMCE 800";
                public override string ChannelF => "852937";
                public override string ChannelR => "852.93750";
                public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
            }
            public class SS_450
            {
                public class CH1 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH2 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH3 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH4 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH5 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH6 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }
                public class CH7 : ASubSS
                {
                    public override string BuildingNameF => "Building: _PRMCE 800";
                    public override string ChannelF => "852937";
                    public override string ChannelR => "852.93750";
                    public override string FloorSectionHeadingTable_Band_CellR => "800 Mhz";
                }

            }
        }
    }
}