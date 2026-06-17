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
        public const string BuildStamp = "p145";

        // p145: GOŁE tokeny dokładnie jak ręczny `-PURGE All * N`, który u usera
        // działa. Wariant z podkreślnikami ("_-PURGE _All * _No") urywał się na
        // "Enter name(s) to purge <*>:" → *Cancel* i purge nie zachodził.
        // Spacja = Enter; KOŃCOWA spacja = ostatni Enter (verify N).
        // -PURGE↵ All↵ *↵ N↵  — 1 przebieg (zweryfikowany ręcznie jako bezpieczny).
        private const string PurgeMacro = "-PURGE All * N ";

        public static void RunNativePurge(Document doc)
        {
            if (doc == null) return;
            doc.SendStringToExecute(PurgeMacro, true, false, false);
        }
    }
}
