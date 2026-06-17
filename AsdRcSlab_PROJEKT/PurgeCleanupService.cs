using Autodesk.AutoCAD.ApplicationServices;

namespace AsdRcSlab
{
    /// <summary>
    /// ASD-PRG (v4, p143): czyszczenie = GOŁA natywna komenda AutoCAD <c>-PURGE</c>
    /// uruchomiona przez <c>doc.SendStringToExecute</c>. Dokładnie to, co bezpieczny
    /// ręczny <c>-PURGE All * N</c> ×3 — i NIC poza tym.
    ///
    /// ZERO managed <c>db.Purge</c>, ZERO <c>Erase</c>, ZERO własnych transakcji /
    /// LockDocument wokół purge. Wcześniejszy ręczny purge + Erase kasował obiekty
    /// referencjonowane przez viewporty/layouty → Access Violation 0x0050 przy
    /// wejściu w layout / plot. Natywne -PURGE respektuje wszystkie referencje.
    /// </summary>
    public static class PurgeCleanupService
    {
        // Stempel wersji — wypisywany na starcie komendy, żeby user wiedział, że
        // załadowała się WŁAŚCIWA (nowa) DLL, a nie stara kopia z innego bundla.
        public const string BuildStamp = "p148";

        // p148: makro = DWA tokeny: "-PURGE" (Enter wywołuje komendę) + "All"
        // (Enter zatwierdza typ → purguje WSZYSTKO i KOŃCZY komendę; All nie pyta
        // o name/verify). Trailing space po "All" to JEDYNY potrzebny Enter.
        // Wcześniejsze "* N" były zbędne: po "All" komenda już się kończyła, więc
        // Enter z "*"/"N" POWTARZAŁ -PURGE i zawieszał na "Enter name(s) to purge
        // <*>:". Wariant z podkreślnikami ("_-PURGE _All …") też się urywał.
        // Głębsze czyszczenie = uruchom komendę ponownie (idempotentne, bezpieczne).
        private const string PurgeMacro = "-PURGE All ";

        public static void RunNativePurge(Document doc)
        {
            if (doc == null) return;
            doc.SendStringToExecute(PurgeMacro, true, false, false);
        }
    }
}
