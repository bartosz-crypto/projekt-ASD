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
        public const string BuildStamp = "p147";

        // p145: GOŁE tokeny dokładnie jak ręczny `-PURGE All * N`, który u usera
        // działa. Wariant z podkreślnikami ("_-PURGE _All * _No") urywał się na
        // "Enter name(s) to purge <*>:" → *Cancel* i purge nie zachodził.
        // Spacja = Enter; KOŃCOWA spacja = ostatni Enter (verify N).
        // p147: POWRÓT do POJEDYNCZEGO przebiegu (wersja wydaniowa). "-PURGE All"
        // czyści natychmiast i NIE pobiera "* N" → przy 3× te tokeny zostawały luzem
        // i rozjeżdżały kolejne przebiegi (drugi -PURGE zawisał na "Enter name(s)…").
        // Głębsze czyszczenie = uruchom komendę ponownie (idempotentne, bezpieczne).
        private const string PurgeMacro = "-PURGE All * N ";

        public static void RunNativePurge(Document doc)
        {
            if (doc == null) return;
            doc.SendStringToExecute(PurgeMacro, true, false, false);
        }
    }
}
