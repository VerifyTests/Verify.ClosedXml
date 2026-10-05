using ClosedXML.Excel;

public class Samples
{
    [Test]
    public Task ScrubbingWithoutFormat() =>
        VerifyFile(ProjectFiles.sample_scrubbingWithoutFormat_xlsx.Path);

    [Test]
    public Task ScrubbingWithoutFormatDisableDateCounting() =>
        VerifyFile(ProjectFiles.sample_scrubbingWithoutFormat_xlsx.Path)
            .DisableDateCounting();

    [Test]
    public Task ScrubbingWithoutFormatDontScrubDateTimes() =>
        VerifyFile(ProjectFiles.sample_scrubbingWithoutFormat_xlsx.Path)
            .DontScrubDateTimes();

    [Test]
    public Task ScrubbingWithoutFormatDontScrubGuids() =>
        VerifyFile(ProjectFiles.sample_scrubbingWithoutFormat_xlsx.Path)
            .DontScrubGuids();

    [Test]
    public Task DontScrub() =>
        VerifyFile(ProjectFiles.sample_xlsx.Path)
            .DontScrubGuids().DontScrubDateTimes();

    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    [Test]
    public Task MultipleSheets() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path);

    [Test]
    public Task HiddenRow() =>
        VerifyFile(ProjectFiles.sample_hidden_row_xlsx.Path);

    #region XLWorkbook

    [Test]
    public Task XLWorkbook()
    {
        using var book = new XLWorkbook();

        var sheet = book.Worksheets.Add("Basic Data");

        sheet.Cell("A1").Value = "ID";
        sheet.Cell("B1").Value = "Name";

        sheet.Cell("A2").Value = 1;
        sheet.Cell("B2").Value = "John Doe";

        sheet.Cell("A3").Value = 2;
        sheet.Cell("B3").Value = "Jane Smith";

        return Verify(book);
    }

    #endregion

    [Test]
    public Task XLWorkbookFromStream()
    {
        using var stream = ProjectFiles.sample_xlsx.OpenRead();
        using var book = new XLWorkbook(stream);
        return Verify(book);
    }

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    #endregion

    #region ExcludeCsv

    [Test]
    public Task ExcludeCsv() =>
        VerifyFile("sample.xlsx")
            .ExcludeDerivedTargets("csv");

    #endregion

    #region ExcludeXlsx

    [Test]
    public Task ExcludeXlsx() =>
        VerifyFile("sample.xlsx")
            .ExcludeTargets("xlsx");

    #endregion

    #region NamedTarget

    [Test]
    public Task NamedTarget()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(new Target("xlsx", stream, "Attachment1"));
    }

    #endregion

    // A hidden sheet is verified as any other, so it has a csv. What says it is hidden is
    // HiddenSheets in the info file.
    [Test]
    public Task HiddenSheet()
    {
        using var book = new XLWorkbook();

        book.Worksheets.Add("Shown").Cell("A1").Value = "first";
        var hidden = book.Worksheets.Add("Hidden");
        hidden.Cell("A1").Value = "second";
        hidden.Visibility = XLWorksheetVisibility.Hidden;

        return Verify(book);
    }

    // A sheet that only code can unhide is verified as one Excel can
    [Test]
    public Task VeryHiddenSheet()
    {
        using var book = new XLWorkbook();

        book.Worksheets.Add("Shown").Cell("A1").Value = "first";
        var hidden = book.Worksheets.Add("Hidden");
        hidden.Cell("A1").Value = "second";
        hidden.Visibility = XLWorksheetVisibility.VeryHidden;

        return Verify(book);
    }

    // A sheet is a page, so PagesToInclude limits the csv files. The info file still names every
    // sheet, and the xlsx is still the whole workbook.
    [Test]
    public Task PagesToInclude() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path)
            .PagesToInclude(_ => _ == 2);

    // A hidden sheet is counted as any other, so here it is the first page
    [Test]
    public Task HiddenSheetIsCountedByPagesToInclude()
    {
        using var book = new XLWorkbook();

        var hidden = book.Worksheets.Add("Hidden");
        hidden.Cell("A1").Value = "first";
        hidden.Visibility = XLWorksheetVisibility.Hidden;
        book.Worksheets.Add("Shown").Cell("A1").Value = "second";

        return Verify(book)
            .PagesToInclude(1);
    }

    [Test]
    public Task SheetNamesWithInvalidFileNameChars()
    {
        using var book = new XLWorkbook();

        book.Worksheets.Add("Q1 <draft>").Cell("A1").Value = "first";
        book.Worksheets.Add("Q1 |draft|").Cell("A1").Value = "second";
        book.Worksheets.Add("Q2 <draft>").Cell("A1").Value = "third";

        return Verify(book);
    }
}