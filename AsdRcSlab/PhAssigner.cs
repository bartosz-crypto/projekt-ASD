using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AsdRcSlab
{
    public static class PhAssigner
    {
        public static List<PileData> AssignAll(List<PileData> piles)
        {
            foreach (var p in piles)
                p.PhAction = AssignOne(p);

            // Grupuj pale wg tego samego PhAction i LocationType -> ApplicablePileIds
            var groups = piles.GroupBy(p => p.PhAction);
            foreach (var grp in groups)
            {
                var ids = grp.Select(p => p.PileId).ToList();
                foreach (var p in grp)
                {
                    p.ApplicablePileIds = ids;
                    p.DetailTitle = GenerateTitle(p, ids);
                }
            }

            return piles;
        }

        private static string AssignOne(PileData p)
        {
            if (p.UtilPct > 100) return "EXCEED";

            switch (p.LocationType.ToUpperInvariant())
            {
                case "INT":
                    if (p.UtilPct == 0)        return "PH1";
                    if (p.UtilPct <= 50)       return "PH4";
                    return "PH7";

                case "EDGE":
                    if (p.UtilPct <= 30)       return "PH2";
                    if (p.UtilPct <= 70)       return "PH5";
                    return "PH8";

                case "CORNER":
                    if (p.UtilPct <= 40)       return "PH3";
                    if (p.UtilPct <= 75)       return "PH6";
                    return "PH9";

                case "REENTRANT":
                    return "PH3-RE";

                default:
                    return "PH1";  // fallback
            }
        }

        public static string GenerateTitle(PileData pile, List<string> groupPileIds)
        {
            // Format: "[LOCATION] PILE CONDITION [PH] EXTRA BARS ([N]No LOCATION[S])
            //          SCALE 1:25 APPLICABLE FOR PILE[S] P01, P02..."

            string locWord = LocationWord(pile.LocationType);
            string phCode  = pile.PhAction;
            int    n       = groupPileIds.Count;

            string pileWord  = n == 1 ? "PILE"     : "PILES";
            string locPlural = n == 1 ? "LOCATION" : "LOCATIONS";

            string pileList = string.Join(", ", groupPileIds.OrderBy(x => x));

            var sb = new StringBuilder();
            sb.Append($"{locWord} PILE CONDITION {phCode}");

            // PH1 nie ma "EXTRA BARS"
            if (phCode == "PH1")
                sb.Append($" (2No LOCATIONS)");
            else if (phCode == "PH3-RE")
                sb.Append($" EXTRA BARS (5No LOCATIONS)");
            else
                sb.Append($" EXTRA BARS ({n}No {locPlural})");

            sb.Append(" SCALE 1:25");
            sb.Append($" APPLICABLE FOR {pileWord} {pileList}");

            return sb.ToString();
        }

        private static string LocationWord(string locType)
        {
            switch (locType.ToUpperInvariant())
            {
                case "INT":       return "INTERNAL";
                case "EDGE":      return "EDGE";
                case "CORNER":    return "CORNER";
                case "REENTRANT": return "REENTRANT";
                default:          return locType.ToUpperInvariant();
            }
        }
    }
}
