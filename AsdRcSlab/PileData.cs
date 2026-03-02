using System.Collections.Generic;

namespace AsdRcSlab
{
    public class PileData
    {
        public string PileId       { get; set; } = "";
        public double UtilPct      { get; set; }
        public string LocationType { get; set; } = "";  // INT / EDGE / CORNER / REENTRANT
        public string PhAction     { get; set; } = "";  // PH1..PH9 / PH3-RE / EXCEED
        public List<string> ApplicablePileIds { get; set; } = new List<string>();
        public string DetailTitle  { get; set; } = "";
    }
}
