namespace VerifyTests;

/// <summary>
/// What <c>Initialize</c> used to be given to choose the outputs a workbook is split into. Settings
/// of Verify choose that now, and nothing reads this: it is here so that code still naming it is
/// told what to use in its place.
/// </summary>
[Obsolete(
    "ClosedXmlOutputs and the outputs argument of Initialize are replaced by settings of Verify: VerifierSettings.ExcludeDerivedTargets(\"csv\") to leave out the csv files. See https://github.com/VerifyTests/Verify.ClosedXml#migrating-from-1x",
    true)]
[Flags]
public enum ClosedXmlOutputs
{
    /// <summary>
    /// No outputs. Only the source document (and info) is emitted.
    /// </summary>
    None = 0,

    /// <summary>
    /// Emit a csv target per worksheet.
    /// </summary>
    Csv = 1,

    /// <summary>
    /// Emit all output kinds.
    /// </summary>
    All = Csv
}
