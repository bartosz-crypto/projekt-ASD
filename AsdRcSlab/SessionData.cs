using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;

namespace AsdRcSlab
{
    /// <summary>
    /// Dane biezacej sesji — wspoldzielone miedzy wszystkimi komendami.
    /// </summary>
    public static class SessionData
    {
        public static ProjectData  CurrentProject { get; set; } = null;
        public static List<PileData> Piles        { get; set; } = new List<PileData>();
        public static bool         PhAssigned     { get; set; } = false;
        public static BmmResult    BmmResults     { get; set; } = null;

        // Lap joint X positions of B1 (horizontal) bars; used by ASD-GTOP to avoid conflicts
        public static List<double> LapPositionsB1 { get; set; } = new List<double>();
        // Lap joint Y positions of B2 (vertical) bars; used by ASD-GTOP to avoid conflicts
        public static List<double> LapPositionsB2 { get; set; } = new List<double>();

        // Template bars registered by ASD-GSETUP (window-selection of pre-drawn ASD bars)
        // Key = bar length rounded to nearest 250 mm; Value = left endpoint of that bar
        public static Dictionary<int, Point3d> TemplateBarsB { get; set; } = new Dictionary<int, Point3d>(); // H10 B1/B2
        public static Dictionary<int, Point3d> TemplateBarsT { get; set; } = new Dictionary<int, Point3d>(); // H12 T1/T2
    }
}
