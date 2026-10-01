public class OutputsTests
{
    [Test]
    public Task ExcludeCsv() =>
        VerifyFile("sample_multiple_sheets.xlsx");
}
