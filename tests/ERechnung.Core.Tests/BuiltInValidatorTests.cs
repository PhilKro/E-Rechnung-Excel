using ERechnung.Core.Calculation;
using ERechnung.Core.Model;
using ERechnung.Core.Samples;
using ERechnung.Core.Sheet;
using ERechnung.Core.Validation;

namespace ERechnung.Core.Tests;

public class BuiltInValidatorTests
{
    private static IEnumerable<string> Codes(Invoice inv, ValidationProfile p = ValidationProfile.XRechnung) =>
        BuiltInValidator.Validate(inv, p).Messages.Where(m => m.Severity == Severity.Error).Select(m => m.Code);

    [Fact]
    public void EmptyInvoice_ReportsCoreAndXRechnungRules()
    {
        var codes = Codes(new Invoice { Payment = new PaymentInstructions { MeansCode = "" } }).ToList();
        foreach (var expected in new[] { "BR-02", "BR-03", "BR-06", "BR-07", "BR-16", "BR-DE-1", "BR-DE-3", "BR-DE-15", "PEPPOL-EN16931-R020" })
            Assert.Contains(expected, codes);
    }

    [Fact]
    public void EN16931Profile_DoesNotRequireXRechnungFields()
    {
        var inv = DemoInvoices.Standard();
        inv.BuyerReference = null;
        inv.Seller.ContactPhone = null;
        Assert.Contains("BR-DE-15", Codes(inv));
        Assert.DoesNotContain("BR-DE-15", Codes(inv, ValidationProfile.EN16931));
        Assert.DoesNotContain("BR-DE-6", Codes(inv, ValidationProfile.EN16931));
    }

    [Fact]
    public void ExemptCategory_RequiresReason()
    {
        var inv = DemoInvoices.Kleinunternehmer();
        inv.VatExemptionReason = null;
        Assert.Contains("BR-E-10", Codes(inv));
    }

    [Fact]
    public void ReverseCharge_RequiresBothVatIds()
    {
        var inv = DemoInvoices.ReverseCharge();
        inv.Buyer.VatId = null;
        Assert.Contains("BR-AE-02", Codes(inv));
    }

    [Fact]
    public void StandardRate_MustBePositive()
    {
        var inv = DemoInvoices.Standard();
        inv.Lines[0].VatRate = 0;
        Assert.Contains("BR-S-05", Codes(inv));
    }

    [Fact]
    public void SellerWithoutVatIdOrIdentifier_FailsBrCo26()
    {
        var inv = DemoInvoices.Kleinunternehmer();
        inv.Seller.Identifier = null;
        Assert.Contains("BR-CO-26", Codes(inv));
    }

    [Theory]
    [InlineData("DE02120300000000202051", true)]
    [InlineData("DE02 1203 0000 0000 2020 51", true)]
    [InlineData("DE02120300000000202052", false)]   // wrong check digits
    [InlineData("DE0212030000000020205", false)]    // too short for DE
    [InlineData("", false)]
    public void Iban_Check(string iban, bool valid) => Assert.Equal(valid, BuiltInValidator.IsValidIban(iban));

    [Fact]
    public void StatedTotals_InconsistentWithLines_AreReported()
    {
        var inv = DemoInvoices.Standard();
        inv.StatedTotals = InvoiceCalculator.Calculate(inv);
        inv.StatedTotals.GrandTotal += 0.01m;
        var codes = Codes(inv).ToList();
        Assert.Contains("BR-CO-15", codes);
        Assert.Contains("BR-CO-16", codes);
    }

    [Fact]
    public void Messages_PointToSheetFields()
    {
        var inv = DemoInvoices.Standard();
        inv.Payment.Iban = null;
        var msg = BuiltInValidator.Validate(inv, ValidationProfile.XRechnung).Messages.Single(m => m.Code == "BR-DE-23-a");
        Assert.Equal(FieldKeys.PaymentIban, msg.FieldKey);
    }

    [Fact]
    public void Calculator_RoundsVatPerGroup_NotPerLine()
    {
        // 10 lines of 0.02 at 19 %: rounding per line would give 10 × round(0.0038) = 0.00,
        // EN 16931 rounds once per VAT group: round(0.20 × 0.19) = round(0.038) = 0.04.
        var inv = new Invoice();
        for (var i = 0; i < 10; i++)
            inv.Lines.Add(new InvoiceLine { Name = "x", Quantity = 1, UnitPrice = 0.02m, VatCategory = "S", VatRate = 19 });

        var t = InvoiceCalculator.Calculate(inv);

        Assert.Equal(0.20m, t.LineTotal);
        Assert.Equal(0.04m, t.TaxTotal);
        Assert.Single(t.VatBreakdown);
    }
}
