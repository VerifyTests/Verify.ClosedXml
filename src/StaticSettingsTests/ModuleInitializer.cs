public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyClosedXml.Initialize();

        // For every test: no csv, so only the workbook and its info are verified
        VerifierSettings.ExcludeDerivedTargets("csv");
    }

    #endregion
}
