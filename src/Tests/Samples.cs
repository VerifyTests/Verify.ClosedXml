using ClosedXML.Excel;

[TestFixture]
public class Samples
{
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
        using var stream = File.OpenRead("sample.xlsx");
        using var reader = new XLWorkbook(stream);
        return Verify(reader);
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    #endregion
}