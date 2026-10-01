public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyClosedXml.Initialize(ClosedXmlOutputs.None);

    #endregion
}
