using ERechnung.Core.Model;

namespace ERechnung.Core.Samples;

/// <summary>Fictitious example invoices: used by the "Beispiel" ribbon button and by the tests.</summary>
public static class DemoInvoices
{
    /// <summary>B2B invoice with 19 % and 7 % lines, SEPA credit transfer. Valid XRechnung 3.0.</summary>
    public static Invoice Standard() => new()
    {
        Number = "RE-2027-0042",
        IssueDate = new DateTime(2027, 3, 15),
        TypeCode = "380",
        CurrencyCode = "EUR",
        DeliveryDate = new DateTime(2027, 3, 10),
        DueDate = new DateTime(2027, 3, 29),
        PaymentTerms = "Zahlbar innerhalb 14 Tagen ohne Abzug.",
        BuyerReference = "04011000-12345-34",
        OrderReference = "PO-7781",
        Note = "Vielen Dank für Ihren Auftrag.",
        Seller = new Party
        {
            Name = "Musterbau GmbH",
            Street = "Hauptstraße 1",
            PostCode = "10115",
            City = "Berlin",
            CountryCode = "DE",
            VatId = "DE123456789",
            TaxNumber = "30/123/45678",
            ElectronicAddress = "rechnung@musterbau.example",
            ContactName = "Erika Mustermann",
            ContactPhone = "+49 30 1234567",
            ContactEmail = "buchhaltung@musterbau.example",
        },
        Buyer = new Party
        {
            Name = "Beispiel Handels AG",
            Identifier = "K-1001",
            Street = "Marktplatz 5",
            PostCode = "80331",
            City = "München",
            CountryCode = "DE",
            VatId = "DE987654321",
            ElectronicAddress = "eingangsrechnungen@beispiel-handel.example",
            ContactName = "Max Beispiel",
        },
        Payment = new PaymentInstructions
        {
            MeansCode = "58",
            Iban = "DE02120300000000202051",
            Bic = "BYLADEM1001",
            AccountName = "Musterbau GmbH",
            RemittanceInformation = "RE-2027-0042",
        },
        Lines =
        {
            new InvoiceLine { Id = "1", ArticleNumber = "MON-01", Name = "Montagearbeiten", Description = "Montage Regalsystem", Quantity = 7.5m, UnitCode = "HUR", UnitPrice = 68.00m, VatCategory = "S", VatRate = 19m },
            new InvoiceLine { Id = "2", ArticleNumber = "REG-200", Name = "Regalsystem Typ 200", Quantity = 3m, UnitCode = "C62", UnitPrice = 249.90m, VatCategory = "S", VatRate = 19m },
            new InvoiceLine { Id = "3", ArticleNumber = "BUCH-1", Name = "Fachbuch Lagerlogistik", Quantity = 2m, UnitCode = "C62", UnitPrice = 39.95m, VatCategory = "S", VatRate = 7m },
        },
    };

    /// <summary>Small business under §19 UStG: no VAT shown, category E with exemption reason.</summary>
    public static Invoice Kleinunternehmer()
    {
        var inv = Standard();
        inv.Number = "KU-2027-007";
        inv.Seller.VatId = null;
        inv.Seller.TaxNumber = "30/987/65432";
        inv.Seller.Identifier = "LIEF-4711";   // BR-CO-26: without a VAT id a seller identifier is required
        inv.VatExemptionReason = "Kein Ausweis von Umsatzsteuer, da Kleinunternehmer gemäß § 19 UStG.";
        foreach (var l in inv.Lines)
        {
            l.VatCategory = "E";
            l.VatRate = 0m;
        }
        return inv;
    }

    /// <summary>Construction service with reverse charge (§13b UStG), category AE.</summary>
    public static Invoice ReverseCharge()
    {
        var inv = Standard();
        inv.Number = "RC-2027-001";
        inv.VatExemptionReason = "Steuerschuldnerschaft des Leistungsempfängers (§ 13b UStG).";
        inv.VatExemptionReasonCode = "VATEX-EU-AE";
        inv.Lines.RemoveAt(2);
        foreach (var l in inv.Lines)
        {
            l.VatCategory = "AE";
            l.VatRate = 0m;
        }
        return inv;
    }
}
