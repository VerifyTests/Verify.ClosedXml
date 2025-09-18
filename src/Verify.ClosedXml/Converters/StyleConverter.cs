class StyleConverter :
    WriteOnlyJsonConverter<IXLStyle>
{
    public override void Write(VerifyJsonWriter writer, IXLStyle target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.Font, "Font");
        if (!target.Alignment.IsDefault())
        {
            writer.WriteMember(target, target.Alignment, "Alignment");
        }

        if (!target.Border.IsDefault())
        {
            writer.WriteMember(target, target.Border, "Border");
        }

        if (!target.Fill.IsDefault())
        {
            writer.WriteMember(target, target.Fill, "Fill");
        }

        writer.WriteMember(target, target.IncludeQuotePrefix, "IncludeQuotePrefix", false);

        if (!target.NumberFormat.IsDefault())
        {
            writer.WriteMember(target, target.NumberFormat, "NumberFormat");
        }

        if (!target.Protection.IsDefault())
        {
            writer.WriteMember(target, target.Protection, "Protection");
        }

        if (!target.NumberFormat.IsDefault())
        {
            writer.WriteMember(target, target.DateFormat, "DateFormat");
        }

        writer.WriteEndObject();
    }
}