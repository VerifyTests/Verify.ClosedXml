class ColorConverter :
    WriteOnlyJsonConverter<XLColor>
{
    public override void Write(VerifyJsonWriter writer, XLColor target) =>
        writer.WriteValue(target.Color.ToString());
}