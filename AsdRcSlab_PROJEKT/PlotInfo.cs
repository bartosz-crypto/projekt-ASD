using System.Collections.Generic;

namespace AsdRcSlab
{
    public class PlotInfo
    {
        public string RawHeader { get; set; }   // surowa zawartość komórki, np. "PLOT 10, 32 (8 piles)"
        public string Label     { get; set; }   // czysty nagłówek bez licznika, np. "PLOT 10, 32" (zachowuje oryginalny separator)

        // Zbiór numerów plotu (uporządkowany, distinct) — NIE zakłada ciągłości.
        // "4-5" -> [4,5]; "10, 32" -> [10,32]; "14-15, 28-29" -> [14,15,28,29].
        public List<int> PlotNumbers { get; set; } = new List<int>();

        public int PileCount      { get; set; }
        public int StartRow       { get; set; }
        public int EndRow         { get; set; }
        public int InternalCount  { get; set; }
        public int EdgeCount      { get; set; }
        public int CornerCount    { get; set; }
        public int ReentrantCount { get; set; }

        public int FirstPlotNumber => (PlotNumbers != null && PlotNumbers.Count > 0) ? PlotNumbers[0] : 0;
        public int LastPlotNumber  => (PlotNumbers != null && PlotNumbers.Count > 0) ? PlotNumbers[PlotNumbers.Count - 1] : 0;

        // backward compat
        public int Number     => FirstPlotNumber;
        public int PlotNumber => FirstPlotNumber;

        public bool IsRange => PlotNumbers != null && PlotNumbers.Count > 1;

        // Zwraca pełny zbiór numerów (NIE rozwija First..Last — to już jest zbiór).
        public List<int> AllPlotNumbers => PlotNumbers ?? new List<int>();

        public string DisplayName => Label ?? RawHeader ?? $"PLOT {FirstPlotNumber} ({PileCount} piles)";

        public override string ToString() => DisplayName;
    }
}
