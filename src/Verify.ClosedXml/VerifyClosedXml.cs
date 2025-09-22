using DeterministicIoPackaging;

namespace VerifyTests;

public static class VerifyClosedXml
{
    static List<JsonConverter> converters =
    [
        new InfoConverter(),
        new ColorConverter(),
        new FontConverter(),
        new StyleConverter(),
        new FillConverter(),
        new BorderConverter(),
        new ProtectionConverter(),
        new NumberFormatConverter(),
        new AlignmentConverter(),
        new WorkbookPropertiesConverter()
    ];

    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => Convert(target, settings));
        VerifierSettings.RegisterFileConverter<XLWorkbook>(Convert);
        VerifierSettings.AddExtraSettings(_ => _.Converters.AddRange(converters));
    }

    static ConversionResult Convert(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = new XLWorkbook(stream);
        return Convert(document, settings);
    }

    static ConversionResult Convert(XLWorkbook book, IReadOnlyDictionary<string, object> settings)
    {
        var sheets = Convert(book).ToList();

        var info = new Info
        {
            SheetNames = sheets.Select(_ => _.Name!).ToList(),
            ColumnWidth = book.ColumnWidth,
            Style = book.Style,
            Properties = book.Properties,
            WorksheetCount = book.Worksheets.Count,
            //Theme = workbook.Theme,
            Use1904DateSystem = book.Use1904DateSystem,
            DefaultFont = book.Style.Font.FontName,
            CalculateMode = book.CalculateMode,
            ShowFormulas = book.ShowFormulas,
            ShowGridLines = book.ShowGridLines,
            ShowRowColHeaders = book.ShowRowColHeaders,
            ShowOutlineSymbols = book.ShowOutlineSymbols,
            ShowZeros = book.ShowZeros,
            ShowRuler = book.ShowRuler,
            ShowWhiteSpace = book.ShowWhiteSpace,
        };

        book.Properties.Created = DeterministicPackage.StableDate;
        book.Properties.Modified = DeterministicPackage.StableDate;
        book.Properties.Author = null;

        using var sourceStream = new MemoryStream();
        book.SaveAs(sourceStream);
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("xlsx", resultStream, performConversion: false)];
        if (sheets.Count == 1)
        {
            var (csv, _) = sheets[0];
            targets.Add(new("csv", csv));
        }
        else
        {
            targets.AddRange(sheets.Select(_ => new Target("csv", _.Csv, _.Name)));
        }

        return new(
            info,
            targets,
            () =>
            {
                book.Dispose();
                return Task.CompletedTask;
            });
    }


    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(XLWorkbook document)
    {
        var current = Counter.Current;
        foreach (var sheet in document.Worksheets)
        {
            var builder = new StringBuilder();

            foreach (var row in sheet.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    var (value, replaceCellValue) = GetCellValue(cell, current);
                    builder.Append(Csv.Escape(value));

                    if (replaceCellValue)
                    {
                        cell.Value = value;
                    }

                    if (cell.FormulaA1.Length > 0)
                    {
                        builder.Append($" ({Csv.Escape(cell.FormulaA1)})");
                    }
                    else if (cell.FormulaR1C1.Length > 0)
                    {
                        builder.Append($" ({Csv.Escape(cell.FormulaR1C1)})");
                    }

                    builder.Append(',');
                }

                builder.Length -= 1;
                builder.AppendLineN();
            }

            yield return (builder, sheet.Name);
        }
    }

    static (string value, bool replaceCellValue) GetCellValue(IXLCell cell, Counter counter)
    {
        if (cell.IsEmpty())
        {
            return (string.Empty, false);
        }

        switch (cell.DataType)
        {
            case XLDataType.Number:
                var value = cell.GetDouble();
                if (cell.Style.NumberFormat.Format.Contains('%'))
                {
                    // Percentage
                    return (value.ToString("P", CultureInfo.InvariantCulture), false);
                }

                return (value.ToString(CultureInfo.InvariantCulture), false);

            case XLDataType.Boolean:
                return (cell.GetBoolean().ToString(), false);

            case XLDataType.DateTime:
                var date = cell.GetDateTime();
                if (counter.TryConvert(date, out var dateResult))
                {
                    return (dateResult, true);
                }

                return (DateFormatter.Convert(date), false);

            case XLDataType.TimeSpan:
                return (cell.GetTimeSpan().ToString(), false);

            case XLDataType.Error:
                return (cell.GetError().ToString(), false);

            case XLDataType.Blank:
                return ("", false);

            default:
                var text = cell.GetText();
                if (counter.TryConvert(text, out var result))
                {
                    return (result, true);
                }

                return (text, false);
        }
    }
}