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
        public const string BuildStamp = "p143";

        // "_"-prefixed = niezależne od języka. "_All"=wszystkie typy nazwanych
        // obiektów, "*"=wszystkie nazwy, "_No"=bez pytań o każdy obiekt.
        // 3 zagnieżdżone przebiegi: usunięcie bloku zwalnia warstwę/linetyp itd.
        private const string PurgeOnce = "_-PURGE _All * _No ";

        public static void RunNativePurge(Document doc)
        {
            if (doc == null) return;
            doc.SendStringToExecute(PurgeOnce + PurgeOnce + PurgeOnce, true, false, false);
        }
    }
}
