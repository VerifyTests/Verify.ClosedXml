public class OutputsTests
{
    [Test]
    public Task ExcludeCsv() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path);
}
