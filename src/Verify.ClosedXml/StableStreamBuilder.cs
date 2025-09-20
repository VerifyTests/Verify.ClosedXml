static class StableStreamBuilder
{
    static DateTime stableDate = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    static DateTimeOffset stableDateOffset = new(stableDate);

    public static Stream Build(XLWorkbook book)
    {
        ForcePropsToBeStable(book);

        var sourceStream = new MemoryStream();
        book.SaveAs(sourceStream);

        var targetStream = new MemoryStream();
        using (var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false))
        using (var targetArchive = new ZipArchive(targetStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var sourceEntry in sourceArchive.Entries)
            {
                DuplicateEntry(sourceEntry, targetArchive);
            }
        }

        targetStream.Position = 0;
        return targetStream;
    }

    static void DuplicateEntry(ZipArchiveEntry sourceEntry, ZipArchive targetArchive)
    {
        if (IsPsmdcp(sourceEntry))
        {
            return;
        }

        using var sourceStream = sourceEntry.Open();
        var targetEntry = targetArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
        targetEntry.LastWriteTime = stableDateOffset;
        using (var targetStream = targetEntry.Open())
        {
            if (IsRels(sourceEntry))
            {
                var xml = XDocument.Load(sourceStream);
                PatchRelsXml(xml);
                xml.Save(targetStream);
            }
            else
            {
                sourceStream.CopyTo(targetStream);
            }
        }
    }

    static void PatchRelsXml(XDocument xml)
    {
        var relationships = xml.Descendants(XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")).ToList();
        var psmdcp = relationships
            .Where(rel =>
            {
                var target = rel.Attribute("Target");
                return target != null &&
                       target.Value.EndsWith(".psmdcp");
            })
            .SingleOrDefault();
        psmdcp?.Remove();

        var workbook = relationships
            .Single(_ => _.Attribute("Target")!.Value.EndsWith("xl/workbook.xml"));
        workbook.Attribute("Id")!.SetValue("VerifyClosedXml");
    }

    static bool IsPsmdcp(ZipArchiveEntry entry) =>
        entry.FullName.StartsWith("package/services/metadata/core-properties/") &&
        entry.Name.EndsWith("psmdcp");

    static bool IsRels(ZipArchiveEntry _) =>
        _.FullName == "_rels/.rels";

    static void ForcePropsToBeStable(XLWorkbook book)
    {
        book.Properties.Created = stableDate;
        book.Properties.Modified = stableDate;
        book.Properties.Author = null;
    }
}