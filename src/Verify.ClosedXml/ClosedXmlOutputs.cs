namespace VerifyTests;

/// <summary>
/// Controls which output kinds a workbook is split into when verified.
/// The source xlsx target and the info target are always emitted.
/// </summary>
[Flags]
public enum ClosedXmlOutputs
{
    /// <summary>
    /// Emit a csv target per worksheet.
    /// </summary>
    Csv = 1,

    /// <summary>
    /// Emit all output kinds.
    /// </summary>
    All = Csv
}
