using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;

namespace ERechnung.Core.IO;

/// <summary>
/// Extracts the embedded invoice XML from a ZUGFeRD / Factur-X / XRechnung-in-PDF file.
/// We need the raw bytes (not a re-serialised descriptor) so the KoSIT validator checks exactly what the sender sent.
/// </summary>
internal static class PdfAttachments
{
    /// <summary>Preferred attachment names in order (ZUGFeRD 2.x / Factur-X, ZUGFeRD 1.0, XRechnung).</summary>
    private static readonly string[] KnownNames =
    {
        "factur-x.xml", "zugferd-invoice.xml", "ZUGFeRD-invoice.xml", "xrechnung.xml",
    };

    public static (string FileName, byte[] Content)? FindInvoiceXml(Stream pdf)
    {
        var all = ReadEmbeddedFiles(pdf);
        if (all.Count == 0) return null;
        foreach (var known in KnownNames)
        {
            var hit = all.FirstOrDefault(a => string.Equals(a.FileName, known, StringComparison.OrdinalIgnoreCase));
            if (hit.Content is not null) return hit;
        }
        var anyXml = all.FirstOrDefault(a => a.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        return anyXml.Content is not null ? anyXml : null;
    }

    public static List<(string FileName, byte[] Content)> ReadEmbeddedFiles(Stream pdf)
    {
        var result = new List<(string, byte[])>();
        PdfDocument doc;
        try
        {
            doc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
        }
        catch (Exception ex)
        {
            throw new InvoiceDataException("Die PDF-Datei kann nicht gelesen werden: " + ex.Message, ex);
        }

        var names = doc.Internals.Catalog.Elements.GetDictionary("/Names");
        var embedded = names?.Elements.GetDictionary("/EmbeddedFiles");
        if (embedded is not null) CollectFromNameTree(embedded, result);

        // Fallback: PDF/A-3 associated files listed directly in the catalog (/AF)
        if (result.Count == 0 && doc.Internals.Catalog.Elements.GetArray("/AF") is { } af)
        {
            foreach (var item in af.Elements)
                if (Deref(item) is PdfDictionary spec) AddFileSpec(spec, null, result);
        }
        return result;
    }

    private static void CollectFromNameTree(PdfDictionary node, List<(string, byte[])> result)
    {
        if (node.Elements.GetArray("/Names") is { } pairs)
        {
            for (var i = 0; i + 1 < pairs.Elements.Count; i += 2)
            {
                var key = (Deref(pairs.Elements[i]) as PdfString)?.Value;
                if (Deref(pairs.Elements[i + 1]) is PdfDictionary spec) AddFileSpec(spec, key, result);
            }
        }
        if (node.Elements.GetArray("/Kids") is { } kids)
        {
            foreach (var kid in kids.Elements)
                if (Deref(kid) is PdfDictionary k) CollectFromNameTree(k, result);
        }
    }

    private static void AddFileSpec(PdfDictionary spec, string? key, List<(string, byte[])> result)
    {
        var name = spec.Elements.GetString("/UF");
        if (string.IsNullOrEmpty(name)) name = spec.Elements.GetString("/F");
        if (string.IsNullOrEmpty(name)) name = key ?? "attachment";

        var ef = spec.Elements.GetDictionary("/EF");
        var fileRef = ef?.Elements["/F"] ?? ef?.Elements["/UF"];
        if (Deref(fileRef) is PdfDictionary fileStream && fileStream.Stream is not null)
            result.Add((name, fileStream.Stream.UnfilteredValue));
    }

    private static PdfItem? Deref(PdfItem? item) => item is PdfReference r ? r.Value : item;
}
