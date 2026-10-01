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

    static ClosedXmlOutputs outputs = ClosedXmlOutputs.All;

    /// <param name="outputs">Which output kinds a workbook is split into. Defaults to <see cref="ClosedXmlOutputs.All"/>.</param>
    public static void Initialize(ClosedXmlOutputs outputs = ClosedXmlOutputs.All)
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;
        VerifyClosedXml.outputs = outputs;

        VerifierSettings.RegisterStreamConverter("xlsx", Convert);
        VerifierSettings.RegisterFileConverter<XLWorkbook>((target, _) => Convert(null, target));
        VerifierSettings.AddExtraSettings(_ => _.Converters.AddRange(converters));
    }

    static ConversionResult Convert(string? targetName, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = new XLWorkbook(stream);
        return Convert(targetName, document);
    }

    static ConversionResult Convert(string? targetName, XLWorkbook book)
    {
        var includeCsv = outputs.HasFlag(ClosedXmlOutputs.Csv);
        var sheets = Convert(book, includeCsv).ToList();

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

        using var sourceStream = new MemoryStream();
        book.SaveAs(sourceStream);
        FixPrefixedDefaultNamespaces(sourceStream);
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("xlsx", resultStream, performConversion: false)];
        if (includeCsv)
        {
            AddCsvTargets(targetName, sheets, targets);
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

    static void AddCsvTargets(string? targetName, List<(StringBuilder Csv, string? Name)> sheets, List<Target> targets)
    {
        if (sheets.Count == 1)
        {
            var (csv, sheetName) = sheets[0];
            string? targetAndSheet;
            if (targetName == null)
            {
                targetAndSheet = sheetName;
            }
            else
            {
                targetAndSheet = $"{targetName}-{sheetName}";
            }

            targets.Add(new("csv", csv, targetAndSheet));
        }
        else
        {
            targets.AddRange(
                sheets.Select(sheet =>
                {
                    string? targetAndSheet;
                    if (targetName == null)
                    {
                        targetAndSheet = sheet.Name;
                    }
                    else
                    {
                        targetAndSheet = $"{targetName}-{sheet.Name}";
                    }

                    return new Target("csv", sheet.Csv, targetAndSheet);
                }));
        }
    }

    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(XLWorkbook document, bool includeCsv)
    {
        var counter = Counter.Current;
        foreach (var sheet in document.Worksheets)
        {
            var builder = new StringBuilder();

            foreach (var row in sheet.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    var (value, replaceCellValue) = GetCellValue(cell, counter);

                    // scrubbed values are written back to the workbook, so this runs even when csv is excluded
                    if (replaceCellValue)
                    {
                        cell.Value = value;
                    }

                    if (!includeCsv)
                    {
                        continue;
                    }

                    builder.Append(Csv.Escape(value));

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

                if (builder.Length > 0)
                {
                    builder.Length -= 1;
                    builder.AppendLineN();
                }
            }

            yield return (builder, sheet.Name);
        }
    }

    static void FixPrefixedDefaultNamespaces(MemoryStream stream)
    {
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
        XNamespace spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            XDocument xml;
            using (var entryStream = entry.Open())
            {
                xml = XDocument.Load(entryStream);
            }

            var root = xml.Root;
            if (root == null)
            {
                continue;
            }

            if (root.Name.Namespace != spreadsheetNs)
            {
                continue;
            }

            var prefix = root.GetPrefixOfNamespace(spreadsheetNs);
            if (string.IsNullOrEmpty(prefix))
            {
                continue;
            }

            // Remove the prefixed namespace declaration and add a default one
            var prefixedAttr = root.Attribute(XNamespace.Xmlns + prefix);
            if (prefixedAttr != null)
            {
                prefixedAttr.Remove();
                root.SetAttributeValue("xmlns", spreadsheetNs.NamespaceName);
            }

            using (var entryStream = entry.Open())
            {
                entryStream.SetLength(0);
                xml.Save(entryStream);
            }
        }

        stream.Position = 0;
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