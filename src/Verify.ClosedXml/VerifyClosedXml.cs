namespace VerifyTests;

public static class VerifyClosedXml
{
    internal static List<JsonConverter> converters =
    [
        new ColorConverter(),
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

    static ConversionResult Convert(XLWorkbook workbook, IReadOnlyDictionary<string, object> settings)
    {
        var sheets = Convert(workbook).ToList();

        var info = new Info
        {
            SheetNames = sheets.Select(_ => _.Name!).ToList(),
            Author = workbook.Author,
            ColumnWidth = workbook.ColumnWidth,
            Style = workbook.Style,
            Properties = workbook.Properties,
            WorksheetCount = workbook.Worksheets.Count,
            Theme = workbook.Theme,
            Use1904DateSystem = workbook.Use1904DateSystem,
            DefaultFont = workbook.Style.Font.FontName,
            CalculateMode = workbook.CalculateMode,
            ShowFormulas = workbook.ShowFormulas,
            ShowGridLines = workbook.ShowGridLines,
            ShowRowColHeaders = workbook.ShowRowColHeaders,
            ShowOutlineSymbols = workbook.ShowOutlineSymbols,
            ShowZeros = workbook.ShowZeros,
            ShowRuler = workbook.ShowRuler,
            ShowWhiteSpace = workbook.ShowWhiteSpace,
        };

        //new("xlsx", CloneToStream(document))
        List<Target> targets = [];
        if (sheets.Count == 1)
        {
            var (csv, _) = sheets[0];
            targets.Add(new("csv", csv));
        }
        else
        {
            targets.AddRange(sheets.Select(_ => new Target("csv", _.Csv, _.Name)));
        }

        return new(info, targets, () =>
        {
            workbook.Dispose();
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
                    builder.Append(EscapeCsvValue(cellValue));
                    builder.Append(',');
                }

                builder.Length -= 1;
                builder.AppendLine();
            }

            yield return (builder, sheet.Name);
        }
    }

    static string GetCellValue(IXLCell cell)
    { if (cell.IsEmpty())
            return string.Empty;

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

    static string EscapeCsvValue(string value)
    {
        // Escape CSV special characters
        if (value.Contains(',') ||
            value.Contains('"') ||
            value.Contains('\n') ||
            value.Contains('\r'))
        {
            // Escape quotes by doubling them
            value = value.Replace("\"", "\"\"");
            // Wrap in quotes
            value = "\"" + value + "\"";
        }

        return value;
    }
}