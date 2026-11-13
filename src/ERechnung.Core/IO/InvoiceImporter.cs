using System.Xml;
using ERechnung.Core.Model;
using ERechnung.Core.Zugferd;
using s2industries.ZUGFeRD;

namespace ERechnung.Core.IO;

/// <summary>Result of reading a received e-invoice file.</summary>
public sealed class ImportedInvoice
{
    public required Invoice Invoice { get; init; }

    /// <summary>The raw invoice XML (for PDFs: the embedded attachment). This is what gets validated.</summary>
    public required byte[] Xml { get; init; }

    /// <summary>Human-readable container/syntax, e.g. "XRechnung (UBL)", "ZUGFeRD/Factur-X PDF (CII)".</summary>
    public required string FormatDescription { get; init; }

    /// <summary>XML syntax of <see cref="Xml"/>.</summary>
    public required XmlSyntax Syntax { get; init; }

    /// <summary>BT-24 Spezifikationskennung (CustomizationID / GuidelineID), e.g. urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0.</summary>
    public string? SpecificationId { get; init; }

    /// <summary>True if the file was a PDF with embedded XML.</summary>
    public bool IsPdf { get; init; }
}

/// <summary>Reads XRechnung (UBL/CII) XML files and ZUGFeRD/Factur-X PDFs.</summary>
public static class InvoiceImporter
{
    public static ImportedInvoice Import(string path)
    {
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); }
        catch (Exception ex) { throw new InvoiceDataException($"Datei kann nicht gelesen werden: {ex.Message}", ex); }
        return Import(bytes);
    }

    public static ImportedInvoice Import(byte[] fileContent)
    {
        var isPdf = fileContent.Length >= 4 && fileContent[0] == '%' && fileContent[1] == 'P' && fileContent[2] == 'D' && fileContent[3] == 'F';
        byte[] xml;
        if (isPdf)
        {
            using var ms = new MemoryStream(fileContent);
            var att = PdfAttachments.FindInvoiceXml(ms)
                ?? throw new InvoiceDataException("Die PDF-Datei enthält keine eingebettete Rechnungs-XML. " +
                                                  "Es handelt sich nicht um eine ZUGFeRD-/Factur-X-Rechnung, sondern um eine reine PDF (keine E-Rechnung im Sinne des §14 UStG).");
            xml = att.Content;
        }
        else
        {
            xml = fileContent;
        }

        var (syntax, specId) = Sniff(xml);

        InvoiceDescriptor d;
        try
        {
            using var xs = new MemoryStream(xml);
            d = InvoiceDescriptor.Load(xs);
        }
        catch (Exception ex)
        {
            throw new InvoiceDataException($"Die Rechnungs-XML kann nicht gelesen werden: {ex.Message}", ex);
        }

        var invoice = DescriptorMapper.FromDescriptor(d);
        var syntaxName = syntax == XmlSyntax.Ubl ? "UBL" : "CII";
        var isXRechnung = specId?.Contains("xrechnung", StringComparison.OrdinalIgnoreCase) == true;
        var format = isPdf
            ? $"ZUGFeRD/Factur-X PDF ({syntaxName}{(isXRechnung ? ", XRechnung-Profil" : "")})"
            : isXRechnung ? $"XRechnung ({syntaxName})" : $"EN 16931 XML ({syntaxName})";

        return new ImportedInvoice
        {
            Invoice = invoice,
            Xml = xml,
            FormatDescription = format,
            Syntax = syntax,
            SpecificationId = specId,
            IsPdf = isPdf,
        };
    }

    /// <summary>Detects the syntax from the root element and reads the specification identifier (BT-24).</summary>
    internal static (XmlSyntax Syntax, string? SpecificationId) Sniff(byte[] xml)
    {
        var doc = new XmlDocument { XmlResolver = null };
        try
        {
            using var ms = new MemoryStream(xml);
            using var reader = XmlReader.Create(ms, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            doc.Load(reader);
        }
        catch (XmlException ex)
        {
            throw new InvoiceDataException($"Die Datei ist keine gültige XML-Datei: {ex.Message}", ex);
        }

        var root = doc.DocumentElement ?? throw new InvoiceDataException("Leere XML-Datei.");
        switch (root.LocalName)
        {
            case "Invoice":
            case "CreditNote":
                return (XmlSyntax.Ubl, FirstChildText(root, "CustomizationID"));
            case "CrossIndustryInvoice":
                var ctx = FirstChild(root, "ExchangedDocumentContext");
                var guideline = ctx is null ? null : FirstChild(ctx, "GuidelineSpecifiedDocumentContextParameter");
                return (XmlSyntax.Cii, guideline is null ? null : FirstChildText(guideline, "ID"));
            default:
                throw new InvoiceDataException(
                    $"Unbekanntes XML-Format (Wurzelelement \"{root.LocalName}\"). Erwartet wird eine XRechnung (UBL oder CII) oder ZUGFeRD/Factur-X.");
        }
    }

    private static XmlElement? FirstChild(XmlElement parent, string localName) =>
        parent.ChildNodes.OfType<XmlElement>().FirstOrDefault(e => e.LocalName == localName);

    private static string? FirstChildText(XmlElement parent, string localName) => FirstChild(parent, localName)?.InnerText?.Trim();
}
