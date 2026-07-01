#if NET48 && DEBUG
[TestFixture]
[Apartment(ApartmentState.STA)]
public class ConvertExcelSnapshots
{
    [Test]
    [Explicit]
    public void Run()
    {
        var directory = ProjectFiles.ProjectDirectory;
        var imageFiles = Directory.EnumerateFiles(directory, "*.png").ToList();
        foreach (var file in imageFiles)
        {
            File.Delete(file);
        }

        ExcelRender.Convert(Path.Combine(directory, "sample.xlsx"));
        var excelFiles = Directory.EnumerateFiles(directory, "*.verified.xlsx").ToList();
        foreach (var file in excelFiles)
        {
            if (Path.GetFileName(file).StartsWith('~'))
            {
                continue;
            }

            if (!Path.GetFileName(file).Contains("Net9"))
            {
                continue;
            }

            ExcelRender.Convert(file);
        }
    }
}
#endif