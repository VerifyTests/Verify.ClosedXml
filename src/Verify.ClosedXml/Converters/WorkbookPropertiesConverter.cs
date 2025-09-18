class WorkbookPropertiesConverter :
    WriteOnlyJsonConverter<XLWorkbookProperties>
{
    public override void Write(VerifyJsonWriter writer, XLWorkbookProperties target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.Title, "Title");
        writer.WriteMember(target, target.Subject, "Subject");
        writer.WriteMember(target, target.Category, "Category");
        writer.WriteMember(target, target.Keywords, "Keywords");
        writer.WriteMember(target, target.Comments, "Comments");
        writer.WriteMember(target, target.Status, "Status");
        writer.WriteMember(target, target.Company, "Company");
        writer.WriteMember(target, target.Manager, "Manager");
        writer.WriteEndObject();
    }
}