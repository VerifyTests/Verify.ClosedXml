#if NET48 && DEBUG

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

[TestFixture]
[Apartment(ApartmentState.STA)]
public class ConvertExcelSnapshots
{
    [Test]
    [Explicit]
    public void Run()
    {
        var directory = AttributeReader.GetProjectDirectory();
        var imageFiles = Directory.EnumerateFiles(directory, "*.png").ToList();
        foreach (var file in imageFiles)
        {
            File.Delete(file);
        }

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

        ExcelRender.Convert(Path.Combine(directory, "sample.xlxs"));
    }
}

public static class ExcelRender
{
    public static void Convert(string excelPath)
    {
        Application? excel = null;
        Workbook? book = null;

        try
        {
            excel = new()
            {
                Visible = true,
                DisplayAlerts = true,
                ScreenUpdating = true
            };

            // Open workbook
            book = excel.Workbooks.Open(excelPath);

            foreach (Worksheet sheet in book.Sheets)
            {
                RenderSheet(excelPath, sheet);
            }
        }
        catch (Exception exception)
        {
            throw new("Failed for: " + excelPath, exception);
        }
        finally
        {
            try
            {
                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    static void RenderSheet(string excelPath, Worksheet sheet)
    {
        try
        {
            var range = sheet.UsedRange;
            sheet.Activate();
            range.Select();
            Clipboard.Clear();
            Thread.Sleep(100);
            range.Copy();
            range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

            Thread.Sleep(100);
            using var image = Clipboard.GetImage()!;
            var imageFile = excelPath.Replace(".verified.xlsx", $"_{sheet.Name}_.png");
            image.Save(imageFile, ImageFormat.Png);
        }
        finally
        {
            Marshal.ReleaseComObject(sheet);
        }
    }
}
#endif