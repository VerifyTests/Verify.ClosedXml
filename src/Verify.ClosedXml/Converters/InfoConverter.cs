class InfoConverter :
    WriteOnlyJsonConverter<Info>
{
    public override void Write(VerifyJsonWriter writer, Info target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.SheetNames, "SheetNames");
        writer.WriteMember(target, target.Properties, "Properties");
        writer.WriteMember(target, target.WorksheetCount, "WorksheetCount");
        writer.WriteMember(target, target.Use1904DateSystem, "Use1904DateSystem", false);
        writer.WriteMember(target, target.DefaultFont, "DefaultFont");
        writer.WriteMember(target, target.CalculateMode, "CalculateMode", XLCalculateMode.Manual);
        writer.WriteMember(target, target.ShowFormulas, "ShowFormulas", false);
        writer.WriteMember(target, target.ShowGridLines, "ShowGridLines", true);
        writer.WriteMember(target, target.ShowRowColHeaders, "ShowRowColHeaders", true);
        writer.WriteMember(target, target.ShowOutlineSymbols, "ShowOutlineSymbols", true);
        writer.WriteMember(target, target.ShowZeros, "ShowZeros", true);
        writer.WriteMember(target, target.ShowRuler, "ShowRuler", true);
        writer.WriteMember(target, target.ShowWhiteSpace, "ShowWhiteSpace", true);
        writer.WriteMember(target, target.Style, "Style");
        writer.WriteEndObject();
    }
}