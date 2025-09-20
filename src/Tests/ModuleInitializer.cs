public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyClosedXml.Initialize();

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifierSettings.UniqueForTargetFrameworkAndVersion();
        VerifierSettings.InitializePlugins();
    }
}