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
}