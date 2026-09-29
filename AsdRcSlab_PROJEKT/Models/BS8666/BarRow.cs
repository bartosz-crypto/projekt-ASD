using System;

namespace AsdRcSlab
{
    /// <summary>
    /// Pojedynczy wiersz z xlsx eksportu zbrojenia (b1.xlsx style):
    /// Bar mark / Type and size / No. members / No. each / Total /
    /// Length per bar / Shape code / A / B / C / D / E or R.
    /// </summary>
    public sealed class BbsBarRow
    {
        public int BarMark { get; set; }
        public string TypeSize { get; set; }   // np. "H12", "H10", "H16"
        public int NoMembers { get; set; }
        public int NoEach { get; set; }
        public int Total { get; set; }
        public double LengthPerBar { get; set; }  // raw z modelu (kolumna F)
        public int ShapeCode { get; set; }        // 0, 11, 12, 13, 21, ..., 98
        public double A { get; set; }
        public double B { get; set; }
        public double C { get; set; }
        public double D { get; set; }
        public double? EOrR { get; set; }   // nullable: nie wszystkie shape codes używają

        /// <summary>
        /// Średnica pręta z TypeSize: "H12" → 12, "H10" → 10, "H16" → 16.
        /// Zakłada prefix jednoznakowy ("H" dla BS).
        /// </summary>
        public int Diameter
        {
            get
            {
                if (string.IsNullOrWhiteSpace(TypeSize) || TypeSize.Length < 2)
                    return 0;
                int d;
                return int.TryParse(TypeSize.Substring(1), out d) ? d : 0;
            }
        }
    }
}
