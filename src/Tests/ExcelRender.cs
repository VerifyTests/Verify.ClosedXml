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

            var imageFile = file.Replace(".verified.xlsx",".png");
            ExcelRender.Convert(file, imageFile);
        }
    }
}

public static class ExcelRender
{
    public static void Convert(string excelPath, string pngPath)
    {
        Application? excel = null;
        Workbook? workbook = null;
        Worksheet? worksheet = null;

        try
        {
            excel = new()
            {
                Visible = true,
                DisplayAlerts = false,
                ScreenUpdating = true
            };

            // Open workbook
            workbook = excel.Workbooks.Open(excelPath);
            worksheet = (Worksheet) workbook.Sheets[1];

            var range = worksheet.UsedRange;

            range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

            Thread.Sleep(1000);
            using var image = Clipboard.GetImage()!;
            image.Save(pngPath, ImageFormat.Png);
        }
        finally
        {
            try
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }

                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

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
}
#endif