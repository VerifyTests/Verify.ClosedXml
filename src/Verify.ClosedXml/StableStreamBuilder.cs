static class StableStreamBuilder
{
    public static Stream Build(XLWorkbook book)
    {
        ForcePropsToBeStable(book);

        var stream = new MemoryStream();
        book.SaveAs(stream);

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            RemovePsmdcp(archive);

            PatchRels(archive);

            archive.FixWriteTime();
        }

        stream.Position = 0;
        return stream;
    }

    static void RemovePsmdcp(ZipArchive archive)
    {
        var entry = archive.Entries.SingleOrDefault(_ =>
            _.FullName.StartsWith("package/services/metadata/core-properties/") &&
            _.Name.EndsWith("psmdcp"));
        entry?.Delete();
    }

    static void PatchRels(ZipArchive archive)
    {
        var rels = archive.Entries.Single(_ => _.FullName == "_rels/.rels");

        var xml = ReadXml(rels);
        var relationships = xml.Descendants(XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")).ToList();
        var psmdcpRelationship = relationships
            .Where(rel =>
            {
                var target = rel.Attribute("Target");
                return target != null &&
                       target.Value.EndsWith(".psmdcp");
            })
            .SingleOrDefault();
        psmdcpRelationship?.Remove();

        var workbookRelationship = relationships
            .Single(_ => _.Attribute("Target")!.Value.EndsWith("xl/workbook.xml"));
        workbookRelationship.Attribute("Id")!.SetValue("Ra40fa3ed832944fe");

        rels.Delete();
        var newRels = archive.CreateEntry("_rels/.rels");
        using var stream = newRels.Open();
        using var streamWriter = new StreamWriter(stream);
        var format = xml.ToString();
        streamWriter.Write(format);
    }

    static XDocument ReadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    static DateTime stableDate = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    static void ForcePropsToBeStable(XLWorkbook book)
    {
        book.Properties.Created = stableDate;
        book.Properties.Modified = stableDate;
        book.Properties.Author = null;
    }
}