namespace AsdRcSlab
{
    /// <summary>
    /// Stan sesji pluginu (in-memory, traci po restarcie AutoCAD).
    /// W przyszłości można rozszerzyć do persystencji w settings.
    /// </summary>
    public static class BbsSessionState
    {
        /// <summary>
        /// Ostatnio użyta ścieżka template BBS — auto-fill w dialogu.
        /// </summary>
        public static string LastTemplatePath { get; set; }
    }
}
