class FillConverter :
    WriteOnlyJsonConverter<IXLFill>
{
    public override void Write(VerifyJsonWriter writer, IXLFill target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.BackgroundColor.Color, "BackgroundColor", System.Drawing.Color.Transparent);
        writer.WriteMember(target, target.PatternColor.Color, "PatternColor", System.Drawing.Color.Transparent);
        writer.WriteMember(target, target.PatternType, "PatternType", XLFillPatternValues.None);
        writer.WriteEndObject();
    }
}