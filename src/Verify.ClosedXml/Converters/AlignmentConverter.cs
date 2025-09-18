class AlignmentConverter :
    WriteOnlyJsonConverter<IXLAlignment>
{
    public override void Write(VerifyJsonWriter writer, IXLAlignment target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.Horizontal, "Horizontal",XLAlignmentHorizontalValues.General);
        writer.WriteMember(target, target.ShrinkToFit, "ShrinkToFit", false);
        writer.WriteMember(target, target.WrapText, "WrapText", false);
        writer.WriteMember(target, target.TopToBottom, "TopToBottom", false);
        writer.WriteEndObject();
    }
}