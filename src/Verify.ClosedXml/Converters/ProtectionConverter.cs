class ProtectionConverter :
    WriteOnlyJsonConverter<IXLProtection>
{
    public override void Write(VerifyJsonWriter writer, IXLProtection target)
    {
        writer.WriteStartObject();
        writer.WriteMember(target, target.Locked, "Locked", true);
        writer.WriteMember(target, target.Hidden, "Hidden", false);
        writer.WriteEndObject();
    }
}