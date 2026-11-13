using ERechnung.Core.IO;
using ERechnung.Core.Model;
using ERechnung.Core.Samples;
using ERechnung.Core.Validation;
using Xunit.Abstractions;

namespace ERechnung.Core.Tests;

/// <summary>
/// Validates generated documents with the official KoSIT validator. These tests need the validator installed via
/// tools/setup-kosit.ps1; without it they pass trivially and say so in the test output.
/// They are the real proof that our output is compliant, so run setup-kosit.ps1 on dev machines.
/// </summary>
public class KositTests
{
    private readonly ITestOutputHelper _out;
    public KositTests(ITestOutputHelper output) => _out = output;

    private KositValidator? Validator()
    {
        var v = KositValidator.TryCreateDefault(out var reason);
        if (v is null) _out.WriteLine($"SKIPPED (KoSIT not installed): {reason}");
        return v;
    }

    public static IEnumerable<object[]> Cases()
    {
        foreach (var syntax in new[] { XmlSyntax.Ubl, XmlSyntax.Cii })
        {
            yield return new object[] { "standard", syntax };
            yield return new object[] { "kleinunternehmer", syntax };
            yield return new object[] { "reverse-charge", syntax };
            yield return new object[] { "creditnote", syntax };
            yield return new object[] { "minimal", syntax };
            yield return new object[] { "lastschrift", syntax };
        }
    }

    internal static Invoice Sample(string name) => name switch
    {
        "standard" => DemoInvoices.Standard(),
        "kleinunternehmer" => DemoInvoices.Kleinunternehmer(),
        "reverse-charge" => DemoInvoices.ReverseCharge(),
        "creditnote" => CreditNote(),
        "minimal" => Minimal(),
        "lastschrift" => DirectDebit(),
        _ => throw new ArgumentOutOfRangeException(nameof(name)),
    };

    /// <summary>Only what XRechnung requires: no terms text (just a due date), no buyer contact, no optional references.</summary>
    private static Invoice Minimal()
    {
        var inv = DemoInvoices.Standard();
        inv.PaymentTerms = null;
        inv.Note = null;
        inv.OrderReference = null;
        inv.Seller.Street = null;
        inv.Seller.TaxNumber = null;
        inv.Buyer.ContactName = null;
        inv.Buyer.Identifier = null;
        inv.Buyer.VatId = null;
        inv.Payment.Bic = null;
        inv.Payment.AccountName = null;
        inv.Payment.RemittanceInformation = null;
        foreach (var l in inv.Lines) { l.ArticleNumber = null; l.Description = null; }
        return inv;
    }

    private static Invoice DirectDebit()
    {
        var inv = DemoInvoices.Standard();
        inv.Payment = new PaymentInstructions
        {
            MeansCode = "59",
            MandateReference = "MANDAT-2025-17",
            CreditorId = "DE98ZZZ09999999999",
            DebitedIban = "DE89370400440532013000",
            RemittanceInformation = "RE-2027-0042",
        };
        inv.PaymentTerms = "Der Betrag wird am 29.03.2027 per SEPA-Lastschrift eingezogen.";
        return inv;
    }

    private static Invoice CreditNote()
    {
        var inv = DemoInvoices.Standard();
        inv.Number = "GS-2027-001";
        inv.TypeCode = "381";
        inv.PrecedingInvoiceNumber = "RE-2027-0042";
        inv.PrecedingInvoiceDate = new DateTime(2027, 3, 15);
        return inv;
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void XRechnung_IsAcceptedByKosit(string sample, XmlSyntax syntax)
    {
        var v = Validator();
        if (v is null) return;

        var xml = InvoiceExporter.ToXRechnungXml(Sample(sample), syntax);
        var result = v.Validate(xml, $"{sample}-{syntax}.xml");

        Dump(result);
        Assert.True(result.Accepted, $"KoSIT rejected {sample}/{syntax}: " +
            string.Join(" | ", result.Messages.Where(m => m.Severity == Severity.Error).Select(m => $"{m.Code}: {m.Text}")));
        Assert.Contains("XRechnung", result.Scenario);
    }

    [Theory]
    [InlineData(ZugferdProfile.EN16931)]
    [InlineData(ZugferdProfile.XRechnung)]
    public void Zugferd_EmbeddedXml_IsAcceptedByKosit(ZugferdProfile profile)
    {
        var v = Validator();
        if (v is null) return;

        // Validate the XML exactly as it sits inside the PDF (extracted again), not just the pre-serialised descriptor.
        using var visual = new MemoryStream(TestFiles.BlankPdf());
        using var pdf = new MemoryStream();
        InvoiceExporter.WriteZugferdPdf(DemoInvoices.Standard(), visual, pdf, profile);
        var embedded = InvoiceImporter.Import(pdf.ToArray()).Xml;
        var result = v.Validate(embedded, $"zugferd-{profile}.xml");

        Dump(result);
        Assert.True(result.Accepted, string.Join(" | ", result.Messages.Select(m => $"{m.Code}: {m.Text}")));
    }

    /// <summary>The built-in pre-check must not report errors for documents KoSIT accepts (no false alarms).</summary>
    [Theory]
    [MemberData(nameof(Cases))]
    public void BuiltInValidator_HasNoErrors_ForValidSamples(string sample, XmlSyntax syntax)
    {
        _ = syntax;
        var report = BuiltInValidator.Validate(Sample(sample), ValidationProfile.XRechnung);
        foreach (var m in report.Messages) _out.WriteLine($"[{m.Severity}] {m.Code}: {m.Text}");
        Assert.Equal(0, report.ErrorCount);
    }

    [Fact]
    public void InvalidInvoice_IsRejected_WithRuleCodes()
    {
        var v = Validator();
        if (v is null) return;

        var inv = DemoInvoices.Standard();
        inv.BuyerReference = null;            // BR-DE-15
        inv.Seller.ContactPhone = null;       // BR-DE-6
        var xml = InvoiceExporter.ToXRechnungXml(inv, XmlSyntax.Ubl);
        var result = v.Validate(xml, "invalid.xml");

        Dump(result);
        Assert.False(result.Accepted);
        Assert.Contains(result.Messages, m => m.Code == "BR-DE-15");
        Assert.NotNull(result.HtmlReport);
    }

    private void Dump(KositResult r)
    {
        _out.WriteLine($"{r.DocumentName}: accepted={r.Accepted} scenario={r.Scenario} html={r.HtmlReport}");
        foreach (var m in r.Messages) _out.WriteLine($"  [{m.Severity}] {m.Code}: {m.Text}  @ {m.Location}");
    }
}
