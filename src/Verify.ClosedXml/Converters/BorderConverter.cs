class BorderConverter :
    WriteOnlyJsonConverter<IXLBorder>
{
    public override void Write(VerifyJsonWriter writer, IXLBorder target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.LeftBorder, "Left", XLBorderStyleValues.None);
        writer.WriteMember(target, target.LeftBorderColor, "LeftColor", XLColor.Black);
        writer.WriteMember(target, target.RightBorder, "Right", XLBorderStyleValues.None);
        writer.WriteMember(target, target.RightBorderColor, "RightColor", XLColor.Black);
        writer.WriteMember(target, target.TopBorder, "Top", XLBorderStyleValues.None);
        writer.WriteMember(target, target.TopBorderColor, "TopColor", XLColor.Black);
        writer.WriteMember(target, target.BottomBorder, "Bottom", XLBorderStyleValues.None);
        writer.WriteMember(target, target.BottomBorderColor, "BottomColor", XLColor.Black);
        writer.WriteMember(target, target.DiagonalBorder, "Diagonal", XLBorderStyleValues.None);
        writer.WriteMember(target, target.DiagonalBorderColor, "DiagonalColor", XLColor.Black);
        writer.WriteMember(target, target.DiagonalUp, "DiagonalUp", false);
        writer.WriteMember(target, target.DiagonalDown, "DiagonalDown", false);
        writer.WriteEndObject();
    }
}