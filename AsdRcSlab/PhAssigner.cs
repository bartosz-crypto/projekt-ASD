using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AsdRcSlab
{
    public static class PhAssigner
    {
        // Tabela: (action, location) → PH
        // ADD H12@200: PH1=INT, PH2=EDGE, PH3=CORNER
        // ADD H16@200: PH4=INT, PH5=EDGE, PH6=CORNER
        // ADD H16@100: PH7=INT, PH8=EDGE, PH9=CORNER
        // REENTRANT  → PH3-RE (niezależnie od action)
        // >100%      → EXCEED

        public static List<PileData> AssignAll(List<PileData> piles)
        {
            foreach (var p in piles)
                p.PhAction = AssignOne(p);

            // Grupuj wg PhAction → ApplicablePileIds + DetailTitle
            foreach (var grp in piles.GroupBy(p => p.PhAction))
            {
                var ids = grp.Select(p => p.PileId).ToList();
                foreach (var p in grp)
                {
                    p.ApplicablePileIds = ids;
                    p.DetailTitle       = GenerateTitle(p, ids);
                }
            }

            return piles;
        }

        private static string AssignOne(PileData p)
        {
            if (p.UtilPct > 100) return "EXCEED";

            string loc    = (p.LocationType ?? "").ToUpperInvariant().Trim();
            string action = (p.PunchingAction ?? "").ToUpperInvariant().Trim();

            if (loc == "REENTRANT") return "PH3-RE";

            // Wyznaczenie poziomu z action
            int level;
            if      (action.Contains("H16") && action.Contains("@100")) level = 3; // ADD H16@100
            else if (action.Contains("H16") && action.Contains("@200")) level = 2; // ADD H16@200
            else                                                          level = 1; // ADD H12@200 / NO ACTION / fallback

            // Mapa level × location → PH number
            //        INT  EDGE  CORNER
            // Lv 1:  PH1  PH2   PH3
            // Lv 2:  PH4  PH5   PH6
            // Lv 3:  PH7  PH8   PH9
            int locIdx;
            switch (loc)
            {
                case "EDGE":   locIdx = 1; break;
                case "CORNER": locIdx = 2; break;
                default:       locIdx = 0; break; // INT
            }

            int phNum = (level - 1) * 3 + locIdx + 1;
            return "PH" + phNum;
        }

        public static string GenerateTitle(PileData pile, List<string> groupPileIds)
        {
            string locWord  = LocationWord(pile.LocationType);
            string phCode   = pile.PhAction;
            int    n        = groupPileIds.Count;
            string pileWord = n == 1 ? "PILE" : "PILES";
            string pileList = string.Join(", ", groupPileIds.OrderBy(x => x));

            var sb = new StringBuilder();
            sb.Append($"{locWord} PILE CONDITION {phCode}");

            if (phCode == "PH3-RE")
                sb.Append($" EXTRA BARS (5No LOCATIONS)");
            else
            {
                string bars = PhToBarDescription(phCode);
                if (!string.IsNullOrEmpty(bars))
                    sb.Append($" EXTRA BARS {bars}");
            }

            sb.Append(" SCALE 1:25");
            sb.Append($" APPLICABLE FOR {pileWord} {pileList}");

            return sb.ToString();
        }

        private static string PhToBarDescription(string ph)
        {
            switch (ph)
            {
                case "PH1": case "PH2": case "PH3": return "H12@200 T1&T2";
                case "PH4": case "PH5": case "PH6": return "H16@200 T1&T2";
                case "PH7": case "PH8": case "PH9": return "H16@100 T1&T2";
                default: return "";
            }
        }

        private static string LocationWord(string locType)
        {
            switch ((locType ?? "").ToUpperInvariant())
            {
                case "INT":       return "INTERNAL";
                case "EDGE":      return "EDGE";
                case "CORNER":    return "CORNER";
                case "REENTRANT": return "REENTRANT";
                default:          return (locType ?? "").ToUpperInvariant();
            }
        }
    }
}
