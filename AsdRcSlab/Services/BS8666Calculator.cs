using System;

namespace AsdRcSlab
{
    /// <summary>
    /// Kalkulator długości pręta wg BS 8666:2020. Bazuje na formułach
    /// z arkusza BS8666_Calculator_v1 (komórka AC16 — mega-IF na shape code).
    /// </summary>
    public static class BS8666Calculator
    {
        /// <summary>
        /// Scheduling radius (promień gięcia) wg BS 8666 Tabela 2.
        /// </summary>
        public static double GetSchedulingRadius(int diameter)
        {
            if (diameter <= 0) return 0;
            if (diameter <= 16) return 2.0 * diameter;
            if (diameter <= 25) return 3.5 * diameter;
            return 4.0 * diameter;
        }

        /// <summary>
        /// Surowa skorygowana długość (z deductions per shape code),
        /// przed zaokrągleniem. Zwraca double.NaN dla shape codes
        /// jeszcze nie zaimplementowanych (p97 to uzupełni).
        /// </summary>
        public static double CalculateRawCuttingLength(BbsBarRow row)
        {
            if (row == null) throw new ArgumentNullException("row");
            int d = row.Diameter;
            double r = GetSchedulingRadius(d);
            double A = row.A, B = row.B, C = row.C, D = row.D;
            double? E = row.EOrR;

            switch (row.ShapeCode)
            {
                case 0:
                    // Straight: L = raw length z modelu (kolumna F)
                    return row.LengthPerBar;

                case 11:
                    // L-shape (90° bend): L = A + B − 0.5r − d
                    return A + B - 0.5 * r - d;

                case 12:
                    // U-shape (2×90°). E/R = actual mandrel radius (user-supplied).
                    // L = A + B − 0.43·R − 1.2·d
                    if (!E.HasValue) return double.NaN;
                    return A + B - 0.43 * E.Value - 1.2 * d;

                case 13:
                    // Crank/offset: L = A + 0.57·B + C − 1.6·d
                    return A + 0.57 * B + C - 1.6 * d;

                case 14:
                    // L = A + C − 4d
                    return A + C - 4 * d;

                case 15:
                    // L = A + C (no deduction)
                    return A + C;

                case 21:
                    // Cranked U-bar (2×90°): L = A + B + C − r − 2d
                    return A + B + C - r - 2 * d;

                case 22:
                    // U-shape double crank (3 bends): L = A + B + C + D − 1.5r − 3d
                    return A + B + C + D - 1.5 * r - 3 * d;

                case 23:
                    // Mega-formula AC16: L = A + B + C − r − 2d (same as code 21)
                    return A + B + C - r - 2 * d;

                case 24:
                    // L = A + B + C
                    return A + B + C;

                case 25:
                    // L = A + B + E. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + E.Value;

                case 26:
                    // L = A + B + C
                    return A + B + C;

                case 27:
                    // L = A + B + C − 0.5r − d
                    return A + B + C - 0.5 * r - d;

                case 28:
                    // L = A + B + C − 0.5r − d
                    return A + B + C - 0.5 * r - d;

                case 29:
                    // L = A + B + C − r − 2d
                    return A + B + C - r - 2 * d;

                case 31:
                    // L = A + B + C + D − 1.5r − 3d
                    return A + B + C + D - 1.5 * r - 3 * d;

                case 32:
                    // L = A + B + C + D − 1.5r − 3d
                    return A + B + C + D - 1.5 * r - 3 * d;

                case 33:
                    // L = 2A + 1.7B + 2C − 4d
                    return 2 * A + 1.7 * B + 2 * C - 4 * d;

                case 34:
                    // L = A + B + C + E − 0.5r − d. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + C + E.Value - 0.5 * r - d;

                case 35:
                    // L = A + B + C + E − 0.5r − d. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + C + E.Value - 0.5 * r - d;

                case 36:
                    // L = A + B + C + D − r − 2d
                    return A + B + C + D - r - 2 * d;

                case 41:
                    // L = A + B + C + D + E − 2r − 4d. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + C + D + E.Value - 2 * r - 4 * d;

                case 44:
                    // L = A + B + C + D + E − 2r − 4d. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + C + D + E.Value - 2 * r - 4 * d;

                case 46:
                    // L = A + 2B + C + E. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + 2 * B + C + E.Value;

                case 47:
                    // L = 2A + B + MAX(21d, 240)
                    return 2 * A + B + Math.Max(21.0 * d, 240.0);

                case 51:
                    // Closed link/stirrup: L = 2A + 2B + MAX(16d, 160)
                    return 2 * A + 2 * B + Math.Max(16.0 * d, 160.0);

                case 56:
                    // L = A + B + C + D + 2E − 2.5r − 5d. E/R required.
                    if (!E.HasValue) return double.NaN;
                    return A + B + C + D + 2 * E.Value - 2.5 * r - 5 * d;

                case 63:
                    // Double rect. link: L = 2A + 3B + MAX(14d, 150)
                    return 2 * A + 3 * B + Math.Max(14.0 * d, 150.0);

                case 67:
                    // L = A
                    return A;

                case 75:
                    // Circular: L = π·(A − d) + B
                    return Math.PI * (A - d) + B;

                case 77:
                    // Helical/spiral. A=outer dia, B=pitch, C=number of turns.
                    // B > A/5 → slanted: L = C · sqrt((π(A−d))² + B²)
                    // else    → flat:    L = C · π(A−d)
                    if (B > A / 5.0)
                        return C * Math.Sqrt(Math.Pow(Math.PI * (A - d), 2) + B * B);
                    return C * Math.PI * (A - d);

                case 98:
                    // L = A + 2B + C + D − 2r − 4d
                    return A + 2 * B + C + D - 2 * r - 4 * d;

                default:
                    return double.NaN;
            }
        }

        /// <summary>
        /// Finalna długość (z round-up do 25 mm dla bent bars).
        /// Straight (code 0) nie podlega zaokrągleniu.
        /// </summary>
        public static double CalculateFinalCuttingLength(BbsBarRow row)
        {
            double raw = CalculateRawCuttingLength(row);
            if (double.IsNaN(raw)) return double.NaN;
            if (row.ShapeCode == 0) return raw;
            return Math.Ceiling(raw / 25.0) * 25.0;
        }
    }
}
