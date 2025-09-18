class FontConverter :
    WriteOnlyJsonConverter<IXLFont>
{
    public override void Write(VerifyJsonWriter writer, IXLFont target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.FontName, "Name");
        writer.WriteMember(target, target.Bold, "Bold", false);
        writer.WriteMember(target, target.Italic, "Italic", false);
        writer.WriteMember(target, target.Underline.ToString(), "Underline", "None");
        writer.WriteMember(target, target.Strikethrough, "Strikethrough", false);
        writer.WriteMember(target, target.Shadow, "Shadow", false);
        writer.WriteMember(target, target.FontCharSet.ToString(), "CharSet", "Default");
        writer.WriteMember(target, target.FontFamilyNumbering.ToString(), "FamilyNumbering", "Swiss");
        writer.WriteEndObject();
    }
}