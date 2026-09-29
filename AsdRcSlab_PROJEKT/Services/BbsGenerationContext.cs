using System;
using System.Collections.Generic;
using System.Linq;

namespace AsdRcSlab
{
    /// <summary>
    /// Pojedynczy layout drawingu z atrybutami title block.
    /// </summary>
    public sealed class BbsLayoutInfo
    {
        public string LayoutName { get; set; }
            // np. "RC010C1" (nazwa layoutu w AutoCAD po rename z ASD-RCN)
        public string DrawingNumber { get; set; }
            // np. "RH149ZS001-RC010" (z atrybutu DRAWING_NUMBER w A1-BL)
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>();
    }

    /// <summary>
    /// Do której warstwy zbrojenia layout odnosi się
    /// (gdy generujemy BBS dla rysunku który pokazuje bottom/top).
    /// </summary>
    public enum BbsLayerAssignment
    {
        Skip         = 0,
        Bottom       = 1,
        Top          = 2,
        BottomAndTop = 3
    }

    /// <summary>
    /// Single layout + user's assignment decision (BOTTOM / TOP / Skip).
    /// </summary>
    public sealed class BbsLayoutAssignment
    {
        public BbsLayoutInfo    Layout     { get; set; }
        public BbsLayerAssignment Assignment { get; set; }
            = BbsLayerAssignment.Skip;
    }

    /// <summary>
    /// Wszystkie dane zebrane przed wygenerowaniem multi-page BBS.
    /// Wypełniane przez user'a w BbsGeneratorDialog.
    /// </summary>
    public sealed class BbsGenerationContext
    {
        public List<BbsLayoutAssignment> Assignments { get; set; }
            = new List<BbsLayoutAssignment>();
        public string ContractNo    { get; set; }
        public string AddressLine1  { get; set; }
        public string AddressLine2  { get; set; }
        public string AddressLine3  { get; set; }
        public string Revision      { get; set; } = "C1";
        public string PlotSuffix    { get; set; }
            // np. "PLOT 1-8" — używane w A3 "REINFORCEMENT DETAILS... PLOT 1-8"

        // p166: Accessories (wiersze 42/43 BBS). Typ z listy + ilość wpisywana z ręki
        // (pusta = zostaw "No." bez liczby, jak w szablonie).
        public string TricTrakType { get; set; } = "TT40";   // TT40 / TT50
        public string TricTrakQty  { get; set; } = "";
        public string HystoolsType { get; set; } = "DK165";  // DK90 / DK165
        public string HystoolsQty  { get; set; } = "";

        // p167: powierzchnia płyty z SLAB NOTES rysunku (null = nie znaleziono).
        public double? SlabAreaM2  { get; set; }

        // p167: ta sama reguła co BS8666_Calculator 'HyStool Calculator' (C19..C21, I39/J39):
        //   net   = area / (2 × spacing), spacing = 1 m (jeden kierunek)
        //   total = ROUNDUP(net × (1 + allowance), 0), allowance = 10 %
        // Obowiązuje dla HyStool i TRIC-TRAK (arkusz liczy obie z tego samego wzoru).
        public const double AccessorySpacingM   = 1.0;
        public const double AccessoryAllowance  = 0.10;

        public static int? SuggestAccessoryQty(double? areaM2)
        {
            if (!areaM2.HasValue || areaM2.Value <= 0) return null;
            double net   = areaM2.Value / (2.0 * AccessorySpacingM);
            double total = net * (1.0 + AccessoryAllowance);
            // Round przed Ceiling: 100/2*1.1 = 55.000000000000007 → bez tego wyszłoby 56
            // (Excel ROUNDUP liczy na 15 cyfrach znaczących, więc daje 55).
            return (int)Math.Ceiling(Math.Round(total, 9));
        }

        /// <summary>"TRIC-TRAK TT40 52No. X 2m" (bez ilości: "TRIC-TRAK TT40 No. X 2m").</summary>
        public string BuildTricTrakLine()
        {
            return "TRIC-TRAK " + (TricTrakType ?? "TT40") + " "
                 + (TricTrakQty ?? "").Trim() + "No. X 2m";
        }

        /// <summary>"CONTINUOUS HYSTOOLS DK165 55No. X 2m" (bez ilości: "... DK165 No. X 2m").</summary>
        public string BuildHystoolsLine()
        {
            return "CONTINUOUS HYSTOOLS " + (HystoolsType ?? "DK165") + " "
                 + (HystoolsQty ?? "").Trim() + "No. X 2m";
        }

        // Computed views
        public List<BbsLayoutInfo> BottomLayouts
        {
            get
            {
                return Assignments
                    .Where(a => a.Assignment == BbsLayerAssignment.Bottom
                             || a.Assignment == BbsLayerAssignment.BottomAndTop)
                    .Select(a => a.Layout)
                    .ToList();
            }
        }

        public List<BbsLayoutInfo> TopLayouts
        {
            get
            {
                return Assignments
                    .Where(a => a.Assignment == BbsLayerAssignment.Top
                             || a.Assignment == BbsLayerAssignment.BottomAndTop)
                    .Select(a => a.Layout)
                    .ToList();
            }
        }

        /// <summary>
        /// Layouts assigned as Bottom + Top combined. Używane przez generator
        /// do decyzji "single sheet with both sections" vs "multi-page".
        /// </summary>
        public List<BbsLayoutInfo> BottomAndTopLayouts
        {
            get
            {
                return Assignments
                    .Where(a => a.Assignment == BbsLayerAssignment.BottomAndTop)
                    .Select(a => a.Layout)
                    .ToList();
            }
        }

        /// <summary>
        /// Wyciąga "plot suffix" z atrybutu TITLE_1 drawingu.
        /// Heurystyka: wszystko przed pierwszą kropką, trim.
        /// Przykłady:
        ///   "PLOT 1-8. REINFORCEMENT DETAILS"      → "PLOT 1-8"
        ///   "HOUSE. REINFORCEMENT DETAILS"         → "HOUSE"
        ///   "REINFORCEMENT DETAILS" (no dot)       → ""
        ///   null/empty                             → ""
        /// </summary>
        private static string ExtractPlotSuffix(string title1)
        {
            if (string.IsNullOrWhiteSpace(title1)) return "";
            string t = title1.Trim();
            int dotIdx = t.IndexOf('.');
            if (dotIdx < 0) return "";  // bez kropki nie pasuje do wzorca
            return t.Substring(0, dotIdx).Trim();
        }

        /// <summary>
        /// Buduje kontekst startowy z listy layoutów drawingu.
        /// Auto-fill heurystyka:
        ///   Contract No. = DRAWING_NUMBER split przez "-" pierwszy człon
        ///   Address      = PROJ_1 / PROJ_2 / PROJ_3
        ///   Plot suffix  = TITLE_1 (wszystko przed pierwszą kropką)
        ///   Revision     = "C1" (default)
        /// </summary>
        public static BbsGenerationContext BuildInitialFromLayouts(
            List<BbsLayoutInfo> layouts)
        {
            var ctx = new BbsGenerationContext();
            foreach (var layout in layouts)
                ctx.Assignments.Add(new BbsLayoutAssignment
                {
                    Layout     = layout,
                    Assignment = BbsLayerAssignment.Skip
                });

            if (layouts.Count > 0)
            {
                var first = layouts[0];

                // Contract No. = DRAWING_NUMBER split przez "-" pierwszy człon
                if (!string.IsNullOrEmpty(first.DrawingNumber))
                {
                    var parts = first.DrawingNumber.Split('-');
                    if (parts.Length > 0) ctx.ContractNo = parts[0];
                }

                // Address = PROJ_1 / PROJ_2 / PROJ_3
                string proj1, proj2, proj3;
                first.Attributes.TryGetValue("PROJ_1", out proj1);
                first.Attributes.TryGetValue("PROJ_2", out proj2);
                first.Attributes.TryGetValue("PROJ_3", out proj3);
                ctx.AddressLine1 = proj1 ?? "";
                ctx.AddressLine2 = proj2 ?? "";
                ctx.AddressLine3 = proj3 ?? "";

                // Plot suffix: weź TITLE_1 pierwszego layoutu
                string title1;
                first.Attributes.TryGetValue("TITLE_1", out title1);
                ctx.PlotSuffix = ExtractPlotSuffix(title1);
            }

            ctx.Revision = "C1";  // default — user może zmienić
            return ctx;
        }
    }
}
