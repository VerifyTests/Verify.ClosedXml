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

        List<Target> targets = [new("xlsx", StableStreamBuilder.Build(book), performConversion: false)];
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
        foreach (var sheet in document.Worksheets)
        {
            var builder = new StringBuilder();

            foreach (var row in sheet.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    var cellValue = GetCellValue(cell);
                    builder.Append(Csv.Escape(cellValue));

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
                builder.AppendLine();
            }

            yield return (builder, sheet.Name);
        }
    }

    static string GetCellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return string.Empty;
        }

        switch (cell.DataType)
        {
            case XLDataType.Text:
                return cell.GetText();

            case XLDataType.Number:
                if (cell.Style.NumberFormat.Format.Contains('%'))
                {
                    // Percentage
                    return cell.GetDouble().ToString("P", CultureInfo.InvariantCulture);
                }

                return cell.GetDouble().ToString(CultureInfo.InvariantCulture);

            case XLDataType.Boolean:
                return cell.GetBoolean().ToString();

            case XLDataType.DateTime:
                var dateTime = cell.GetDateTime();
                return DateFormatter.Convert(dateTime);

            case XLDataType.TimeSpan:
                return cell.GetTimeSpan().ToString();

            case XLDataType.Error:
                return cell.GetError().ToString();

            default:
                return cell.GetText();
        }
    }

}