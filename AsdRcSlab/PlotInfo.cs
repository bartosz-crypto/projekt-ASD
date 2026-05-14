namespace AsdRcSlab
{
    public class PlotInfo
    {
        public int Number;
        public int PileCount;
        public int StartRow;
        public int EndRow;
        public int InternalCount;
        public int EdgeCount;
        public int CornerCount;
        public int ReentrantCount;

        public override string ToString() => $"PLOT {Number} ({PileCount} piles)";
    }
}
