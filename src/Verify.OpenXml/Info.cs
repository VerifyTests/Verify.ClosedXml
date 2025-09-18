class Info
{
    public required IReadOnlyList<string> SheetNames { get; init; }
    public required string Author { get; init; }
    public required double ColumnWidth { get; init; }
    public required XLWorkbookProperties Properties { get; init; }
    public required int WorksheetCount { get; init; }
    public required IXLTheme Theme { get; init; }
    public required bool Use1904DateSystem { get; init; }
    public required string DefaultFont { get; init; }
    public required XLCalculateMode CalculateMode { get; init; }
    public required bool ShowFormulas { get; init; }
    public required bool ShowGridLines { get; init; }
    public required bool ShowRowColHeaders { get; init; }
    public required bool ShowOutlineSymbols { get; init; }
    public required bool ShowZeros { get; init; }
    public required bool ShowRuler { get; init; }
    public required bool ShowWhiteSpace { get; init; }
    public required IXLStyle Style { get; init; }
}