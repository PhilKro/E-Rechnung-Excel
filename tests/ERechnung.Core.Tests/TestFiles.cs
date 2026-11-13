using PdfSharp.Pdf;

namespace ERechnung.Core.Tests;

internal static class TestFiles
{
    /// <summary>Folder for files produced by tests (inspect manually or feed to KoSIT).</summary>
    public static string OutputDir
    {
        get
        {
            var dir = Path.Combine(AppContext.BaseDirectory, "out");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    public static string Save(string fileName, byte[] content)
    {
        var path = Path.Combine(OutputDir, fileName);
        File.WriteAllBytes(path, content);
        return path;
    }

    /// <summary>A minimal one-page PDF that stands in for the Excel-rendered invoice.</summary>
    public static byte[] BlankPdf()
    {
        var doc = new PdfDocument();
        doc.AddPage();
        using var ms = new MemoryStream();
        doc.Save(ms, false);
        return ms.ToArray();
    }
}
