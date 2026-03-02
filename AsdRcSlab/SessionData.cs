using System.Collections.Generic;

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
    }
}
