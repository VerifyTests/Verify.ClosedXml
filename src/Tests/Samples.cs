using ClosedXML.Excel;

[TestFixture]
public class Samples
{
    [Test]
    public Task ScrubbingWithoutFormat() =>
        VerifyFile("sample_scrubbingWithoutFormat.xlsx");

    [Test]
    public Task ScrubbingWithoutFormatDisableDateCounting() =>
        VerifyFile("sample_scrubbingWithoutFormat.xlsx")
            .DisableDateCounting();

    [Test]
    public Task ScrubbingWithoutFormatDontScrubDateTimes() =>
        VerifyFile("sample_scrubbingWithoutFormat.xlsx")
            .DontScrubDateTimes();

    [Test]
    public Task ScrubbingWithoutFormatDontScrubGuids() =>
        VerifyFile("sample_scrubbingWithoutFormat.xlsx")
            .DontScrubGuids();

    [Test]
    public Task DontScrub() =>
        VerifyFile("sample.xlsx")
            .DontScrubGuids().DontScrubDateTimes();


    [Test]
    public Task VerifyExcelArchive() =>
        VerifyZip("sample.xlsx");

    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    [Test]
    public Task MultipleSheets() =>
        VerifyFile("sample_multiple_sheets.xlsx");

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
        using var stream = File.OpenRead("sample.xlsx");
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