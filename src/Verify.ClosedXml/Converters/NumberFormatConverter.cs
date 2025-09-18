class NumberFormatConverter :
    WriteOnlyJsonConverter<IXLNumberFormat>
{
    public override void Write(VerifyJsonWriter writer, IXLNumberFormat target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.NumberFormatId, "Id", 0);
        writer.WriteMember(target, target.Format, "Format", string.Empty);
        writer.WriteEndObject();
    }
}