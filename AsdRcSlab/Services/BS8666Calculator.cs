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
            double E = row.EOrR ?? 0.0;

            switch (row.ShapeCode)
            {
                case 0:
                    // Straight: L = raw length z modelu (kolumna F)
                    return row.LengthPerBar;

                case 11:
                    // L-shape (90° bend): L = A + B − 0.5r − d
                    return A + B - 0.5 * r - d;

                case 21:
                    // Cranked U-bar (2×90°): L = A + B + C − r − 2d
                    return A + B + C - r - 2 * d;

                // p97: codes 12, 13, 14, 15, 22, 23, 24, 25, 26, 27, 28,
                // 29, 31, 32, 33, 34, 35, 36, 41, 44, 46, 47, 51, 56, 63,
                // 67, 75, 77, 98 (28 sztuk).
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
