using ERechnung.Core.Model;
using ERechnung.Core.Zugferd;
using s2industries.ZUGFeRD;
using s2industries.ZUGFeRD.PDF;

namespace ERechnung.Core.IO;

/// <summary>XML syntax of an XRechnung.</summary>
public enum XmlSyntax
{
    /// <summary>OASIS UBL 2.1 (Invoice / CreditNote).</summary>
    Ubl,

    /// <summary>UN/CEFACT Cross Industry Invoice (CII) D16B.</summary>
    Cii,
}

/// <summary>Profile of the XML embedded in a ZUGFeRD / Factur-X PDF.</summary>
public enum ZugferdProfile
{
    /// <summary>ZUGFeRD 2.3 / Factur-X profile EN 16931 ("Comfort"). Standard choice for B2B.</summary>
    EN16931,

    /// <summary>ZUGFeRD 2.3 with the XRechnung profile (for public-sector recipients who accept hybrid PDFs).</summary>
    XRechnung,
}

/// <summary>Writes invoices as XRechnung XML or ZUGFeRD/Factur-X hybrid PDF.</summary>
public static class InvoiceExporter
{
    /// <summary>XRechnung 3.0 as UBL or CII.</summary>
    public static void WriteXRechnung(Invoice invoice, Stream output, XmlSyntax syntax)
    {
        var d = DescriptorMapper.ToDescriptor(invoice);
        d.Save(output, ZUGFeRDVersion.Version23, Profile.XRechnung,
            syntax == XmlSyntax.Ubl ? ZUGFeRDFormats.UBL : ZUGFeRDFormats.CII);
    }

    public static byte[] ToXRechnungXml(Invoice invoice, XmlSyntax syntax)
    {
        using var ms = new MemoryStream();
        WriteXRechnung(invoice, ms, syntax);
        return ms.ToArray();
    }

    /// <summary>The CII XML that <see cref="WriteZugferdPdf"/> embeds (useful for validating before writing the PDF).</summary>
    public static byte[] ToZugferdXml(Invoice invoice, ZugferdProfile profile)
    {
        var d = DescriptorMapper.ToDescriptor(invoice);
        using var ms = new MemoryStream();
        d.Save(ms, ZUGFeRDVersion.Version23, ToLibraryProfile(profile), ZUGFeRDFormats.CII);
        return ms.ToArray();
    }

    /// <summary>
    /// Creates a ZUGFeRD 2.3 / Factur-X PDF: takes an existing PDF with the visual invoice
    /// (e.g. exported from the Excel sheet), embeds the CII XML and adds PDF/A-3 metadata.
    /// </summary>
    public static void WriteZugferdPdf(Invoice invoice, Stream visualPdf, Stream output, ZugferdProfile profile = ZugferdProfile.EN16931)
    {
        var d = DescriptorMapper.ToDescriptor(invoice);
        InvoicePdfProcessor.SaveToPdf(output, ZUGFeRDVersion.Version23, ToLibraryProfile(profile),
            ZUGFeRDFormats.CII, visualPdf, d, null);
    }

    private static Profile ToLibraryProfile(ZugferdProfile p) =>
        p == ZugferdProfile.XRechnung ? Profile.XRechnung : Profile.Comfort;
}
