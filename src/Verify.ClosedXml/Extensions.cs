static class Extensions
{
    static DateTimeOffset stableDate = new(new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc));

    public static void FixWriteTime(this ZipArchive archive)
    {
        foreach (var entry in archive.Entries)
        {
            entry.LastWriteTime = stableDate;
        }
    }
}