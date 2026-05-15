using System.Collections.Generic;

namespace AsdRcSlab
{
    public class PlotInfo
    {
        public string RawHeader         { get; set; }
        public int    FirstPlotNumber   { get; set; }
        public int    LastPlotNumber    { get; set; }
        public int    PileCount         { get; set; }
        public int    StartRow          { get; set; }
        public int    EndRow            { get; set; }
        public int    InternalCount     { get; set; }
        public int    EdgeCount         { get; set; }
        public int    CornerCount       { get; set; }
        public int    ReentrantCount    { get; set; }

        // backward compat
        public int Number
        {
            get => FirstPlotNumber;
            set => FirstPlotNumber = value;
        }
        public int PlotNumber => FirstPlotNumber;

        public bool IsRange => LastPlotNumber > FirstPlotNumber;

        public List<int> AllPlotNumbers
        {
            get
            {
                var list = new List<int>();
                for (int n = FirstPlotNumber; n <= LastPlotNumber; n++) list.Add(n);
                return list;
            }
        }

        public string DisplayName => RawHeader ?? $"PLOT {FirstPlotNumber} ({PileCount} piles)";

        public override string ToString() => DisplayName;
    }
}
