using ERechnung.Core.Calculation;
using ERechnung.Core.IO;
using ERechnung.Core.Model;
using ERechnung.Core.Samples;

namespace ERechnung.Core.Tests;

public class ExportImportTests
{
    [Theory]
    [InlineData(XmlSyntax.Ubl)]
    [InlineData(XmlSyntax.Cii)]
    public void XRechnung_RoundTrip_PreservesContent(XmlSyntax syntax)
    {
        var original = DemoInvoices.Standard();
        var xml = InvoiceExporter.ToXRechnungXml(original, syntax);
        TestFiles.Save($"standard-{syntax}.xml", xml);

        var imported = InvoiceImporter.Import(xml);

        Assert.Equal(syntax, imported.Syntax);
        Assert.False(imported.IsPdf);
        Assert.Contains("xrechnung", imported.SpecificationId, StringComparison.OrdinalIgnoreCase);
        AssertSameContent(original, imported.Invoice);
    }

    [Fact]
    public void ZugferdPdf_RoundTrip_ExtractsEmbeddedXml()
    {
        var original = DemoInvoices.Standard();
        using var visual = new MemoryStream(TestFiles.BlankPdf());
        using var output = new MemoryStream();

        InvoiceExporter.WriteZugferdPdf(original, visual, output, ZugferdProfile.EN16931);
        var pdf = output.ToArray();
        TestFiles.Save("standard-zugferd.pdf", pdf);
        TestFiles.Save("standard-zugferd-embedded.xml", InvoiceExporter.ToZugferdXml(original, ZugferdProfile.EN16931));

        var imported = InvoiceImporter.Import(pdf);

        Assert.True(imported.IsPdf);
        Assert.Equal(XmlSyntax.Cii, imported.Syntax);
        Assert.Contains("en16931", imported.SpecificationId, StringComparison.OrdinalIgnoreCase);
        AssertSameContent(original, imported.Invoice);
    }

    [Fact]
    public void PlainPdf_WithoutXml_IsRejectedWithClearMessage()
    {
        var ex = Assert.Throws<InvoiceDataException>(() => InvoiceImporter.Import(TestFiles.BlankPdf()));
        Assert.Contains("keine eingebettete Rechnungs-XML", ex.Message);
    }

    [Fact]
    public void UnknownXml_IsRejected()
    {
        var ex = Assert.Throws<InvoiceDataException>(() => InvoiceImporter.Import("<foo/>"u8.ToArray()));
        Assert.Contains("Unbekanntes XML-Format", ex.Message);
    }

    [Theory]
    [InlineData("kleinunternehmer")]
    [InlineData("reverse-charge")]
    public void SpecialVatCases_RoundTrip(string which)
    {
        var original = which == "kleinunternehmer" ? DemoInvoices.Kleinunternehmer() : DemoInvoices.ReverseCharge();
        foreach (var syntax in new[] { XmlSyntax.Ubl, XmlSyntax.Cii })
        {
            var xml = InvoiceExporter.ToXRechnungXml(original, syntax);
            TestFiles.Save($"{which}-{syntax}.xml", xml);
            var imported = InvoiceImporter.Import(xml).Invoice;
            AssertSameContent(original, imported);
            Assert.Equal(original.VatExemptionReason, imported.VatExemptionReason);
        }
    }

    [Fact]
    public void CreditNote_Ubl_RoundTrip()
    {
        var original = DemoInvoices.Standard();
        original.Number = "GS-2027-001";
        original.TypeCode = "381";
        original.PrecedingInvoiceNumber = "RE-2027-0042";
        original.PrecedingInvoiceDate = new DateTime(2027, 3, 15);
        var xml = InvoiceExporter.ToXRechnungXml(original, XmlSyntax.Ubl);
        TestFiles.Save("creditnote-Ubl.xml", xml);

        var imported = InvoiceImporter.Import(xml).Invoice;

        Assert.Equal("381", imported.TypeCode);
        Assert.Equal("RE-2027-0042", imported.PrecedingInvoiceNumber);
        // Known limitation of ZUGFeRD-csharp 18: a UBL CreditNote carries no due date (BT-9). See AGENTS.md.
        Assert.Null(imported.DueDate);
        original.DueDate = null;
        AssertSameContent(original, imported);
    }

    private static void AssertSameContent(Invoice expected, Invoice actual)
    {
        Assert.Equal(expected.Number, actual.Number);
        Assert.Equal(expected.IssueDate, actual.IssueDate);
        Assert.Equal(expected.TypeCode, actual.TypeCode);
        Assert.Equal(expected.CurrencyCode, actual.CurrencyCode);
        Assert.Equal(expected.DueDate, actual.DueDate);
        Assert.Equal(expected.DeliveryDate, actual.DeliveryDate);
        Assert.Equal(expected.BuyerReference, actual.BuyerReference);
        Assert.Equal(expected.OrderReference, actual.OrderReference);

        Assert.Equal(expected.Seller.Name, actual.Seller.Name);
        Assert.Equal(expected.Seller.Street, actual.Seller.Street);
        Assert.Equal(expected.Seller.PostCode, actual.Seller.PostCode);
        Assert.Equal(expected.Seller.City, actual.Seller.City);
        Assert.Equal(expected.Seller.CountryCode, actual.Seller.CountryCode);
        Assert.Equal(expected.Seller.VatId, actual.Seller.VatId);
        Assert.Equal(expected.Seller.TaxNumber, actual.Seller.TaxNumber);
        Assert.Equal(expected.Seller.ElectronicAddress, actual.Seller.ElectronicAddress);
        Assert.Equal(expected.Seller.ContactName, actual.Seller.ContactName);
        Assert.Equal(expected.Seller.ContactPhone, actual.Seller.ContactPhone);
        Assert.Equal(expected.Seller.ContactEmail, actual.Seller.ContactEmail);

        Assert.Equal(expected.Buyer.Name, actual.Buyer.Name);
        Assert.Equal(expected.Buyer.City, actual.Buyer.City);
        Assert.Equal(expected.Buyer.VatId, actual.Buyer.VatId);
        Assert.Equal(expected.Buyer.ElectronicAddress, actual.Buyer.ElectronicAddress);

        Assert.Equal(expected.Payment.MeansCode, actual.Payment.MeansCode);
        Assert.Equal(expected.Payment.Iban, actual.Payment.Iban);
        Assert.Equal(expected.Payment.RemittanceInformation, actual.Payment.RemittanceInformation);

        Assert.Equal(expected.Lines.Count, actual.Lines.Count);
        for (var i = 0; i < expected.Lines.Count; i++)
        {
            var e = expected.Lines[i];
            var a = actual.Lines[i];
            Assert.Equal(e.Id, a.Id);
            Assert.Equal(e.Name, a.Name);
            Assert.Equal(e.ArticleNumber, a.ArticleNumber);
            Assert.Equal(e.Quantity, a.Quantity);
            Assert.Equal(e.UnitCode, a.UnitCode);
            Assert.Equal(e.UnitPrice, a.UnitPrice);
            Assert.Equal(e.VatCategory, a.VatCategory);
            Assert.Equal(e.VatRate, a.VatRate);
            Assert.Equal(InvoiceCalculator.LineNet(e), InvoiceCalculator.LineNet(a));
        }

        var expectedTotals = InvoiceCalculator.Calculate(expected);
        Assert.NotNull(actual.StatedTotals);
        Assert.Equal(expectedTotals.LineTotal, actual.StatedTotals!.LineTotal);
        Assert.Equal(expectedTotals.TaxTotal, actual.StatedTotals.TaxTotal);
        Assert.Equal(expectedTotals.GrandTotal, actual.StatedTotals.GrandTotal);
        Assert.Equal(expectedTotals.DuePayable, actual.StatedTotals.DuePayable);
    }
}
