namespace AsdRcSlab
{
    /// <summary>
    /// Dane biezacej sesji — wspoldzielone miedzy wszystkimi komendami.
    /// </summary>
    public static class SessionData
    {
        public static ProjectData CurrentProject { get; set; } = null;
    }
}
