namespace ERechnung.Core.Model;

/// <summary>
/// Neutral invoice model covering the EN 16931 subset this add-in supports.
/// Business term numbers (BT-x / BG-x) refer to EN 16931-1.
/// This is the single model shared by the Excel sheet, the export, the import and the validators.
/// </summary>
public sealed class Invoice
{
    /// <summary>BT-1 Rechnungsnummer.</summary>
    public string? Number { get; set; }

    /// <summary>BT-2 Rechnungsdatum.</summary>
    public DateTime? IssueDate { get; set; }

    /// <summary>BT-3 Rechnungsart (UNTDID 1001), e.g. 380 = Rechnung, 381 = Gutschrift.</summary>
    public string TypeCode { get; set; } = "380";

    /// <summary>BT-5 Währung (ISO 4217).</summary>
    public string CurrencyCode { get; set; } = "EUR";

    /// <summary>BT-9 Fälligkeitsdatum.</summary>
    public DateTime? DueDate { get; set; }

    /// <summary>BT-72 Liefer-/Leistungsdatum.</summary>
    public DateTime? DeliveryDate { get; set; }

    /// <summary>BT-73 Leistungszeitraum Beginn.</summary>
    public DateTime? PeriodStart { get; set; }

    /// <summary>BT-74 Leistungszeitraum Ende.</summary>
    public DateTime? PeriodEnd { get; set; }

    /// <summary>BT-10 Käuferreferenz, bei öffentlichen Auftraggebern die Leitweg-ID.</summary>
    public string? BuyerReference { get; set; }

    /// <summary>BT-13 Bestellnummer.</summary>
    public string? OrderReference { get; set; }

    /// <summary>BT-12 Vertragsnummer.</summary>
    public string? ContractReference { get; set; }

    /// <summary>BT-25 Nummer der vorausgegangenen Rechnung (bei Korrektur/Gutschrift).</summary>
    public string? PrecedingInvoiceNumber { get; set; }

    /// <summary>BT-26 Datum der vorausgegangenen Rechnung.</summary>
    public DateTime? PrecedingInvoiceDate { get; set; }

    /// <summary>BT-22 Bemerkung zur Rechnung.</summary>
    public string? Note { get; set; }

    /// <summary>BT-20 Zahlungsbedingungen (Freitext).</summary>
    public string? PaymentTerms { get; set; }

    /// <summary>BT-120 Grund der Steuerbefreiung (gilt für alle Positionen mit Kategorie E, AE, K, G, O).</summary>
    public string? VatExemptionReason { get; set; }

    /// <summary>BT-121 Code für den Befreiungsgrund (VATEX-Liste), optional.</summary>
    public string? VatExemptionReasonCode { get; set; }

    /// <summary>BT-113 Bereits gezahlter Betrag (Anzahlungen).</summary>
    public decimal PrepaidAmount { get; set; }

    /// <summary>BG-4 Verkäufer.</summary>
    public Party Seller { get; set; } = new();

    /// <summary>BG-7 Käufer.</summary>
    public Party Buyer { get; set; } = new();

    /// <summary>BG-16 Zahlungsanweisungen.</summary>
    public PaymentInstructions Payment { get; set; } = new();

    /// <summary>BG-25 Rechnungspositionen.</summary>
    public List<InvoiceLine> Lines { get; set; } = new();

    /// <summary>
    /// Totals as stated in a received invoice (BG-22/BG-23). <c>null</c> for invoices created in Excel:
    /// those are always computed by <see cref="Calculation.InvoiceCalculator"/>.
    /// </summary>
    public InvoiceTotals? StatedTotals { get; set; }
}

/// <summary>Seller (BG-4) or buyer (BG-7).</summary>
public sealed class Party
{
    /// <summary>BT-27 / BT-44 Name.</summary>
    public string? Name { get; set; }

    /// <summary>BT-29 / BT-46 Kennung (Lieferanten-/Kundennummer).</summary>
    public string? Identifier { get; set; }

    /// <summary>BT-35 / BT-50 Straße und Hausnummer.</summary>
    public string? Street { get; set; }

    /// <summary>BT-38 / BT-53 Postleitzahl.</summary>
    public string? PostCode { get; set; }

    /// <summary>BT-37 / BT-52 Ort.</summary>
    public string? City { get; set; }

    /// <summary>BT-40 / BT-55 Ländercode (ISO 3166-1 alpha-2).</summary>
    public string CountryCode { get; set; } = "DE";

    /// <summary>BT-31 / BT-48 Umsatzsteuer-Identifikationsnummer.</summary>
    public string? VatId { get; set; }

    /// <summary>BT-32 Steuernummer (nur Verkäufer).</summary>
    public string? TaxNumber { get; set; }

    /// <summary>BT-34 / BT-49 Elektronische Adresse (z. B. E-Mail, Leitweg-ID, Peppol-ID).</summary>
    public string? ElectronicAddress { get; set; }

    /// <summary>
    /// Schema der elektronischen Adresse (CEF EAS), z. B. EM = E-Mail, 0204 = Leitweg-ID, 9930 = deutsche USt-IdNr.
    /// Leer = automatisch erkennen (siehe <see cref="Codes.ElectronicAddressSchemes.Detect"/>).
    /// </summary>
    public string? ElectronicAddressScheme { get; set; }

    /// <summary>BT-41 / BT-56 Ansprechpartner.</summary>
    public string? ContactName { get; set; }

    /// <summary>BT-42 / BT-57 Telefon.</summary>
    public string? ContactPhone { get; set; }

    /// <summary>BT-43 / BT-58 E-Mail des Ansprechpartners.</summary>
    public string? ContactEmail { get; set; }
}

/// <summary>BG-16 Zahlungsanweisungen.</summary>
public sealed class PaymentInstructions
{
    /// <summary>BT-81 Zahlungsart (UNTDID 4461), z. B. 58 = SEPA-Überweisung, 59 = SEPA-Lastschrift.</summary>
    public string MeansCode { get; set; } = "58";

    /// <summary>BT-83 Verwendungszweck.</summary>
    public string? RemittanceInformation { get; set; }

    /// <summary>BT-84 IBAN des Zahlungsempfängers (Überweisung).</summary>
    public string? Iban { get; set; }

    /// <summary>BT-86 BIC des Zahlungsempfängers.</summary>
    public string? Bic { get; set; }

    /// <summary>BT-85 Kontoinhaber.</summary>
    public string? AccountName { get; set; }

    /// <summary>BT-89 Mandatsreferenz (Lastschrift).</summary>
    public string? MandateReference { get; set; }

    /// <summary>BT-90 Gläubiger-Identifikationsnummer (Lastschrift).</summary>
    public string? CreditorId { get; set; }

    /// <summary>BT-91 IBAN des belasteten Kontos (Lastschrift).</summary>
    public string? DebitedIban { get; set; }
}

/// <summary>BG-25 Rechnungsposition.</summary>
public sealed class InvoiceLine
{
    /// <summary>BT-126 Positionsnummer.</summary>
    public string? Id { get; set; }

    /// <summary>BT-155 Artikelnummer des Verkäufers.</summary>
    public string? ArticleNumber { get; set; }

    /// <summary>BT-153 Bezeichnung.</summary>
    public string? Name { get; set; }

    /// <summary>BT-154 Beschreibung.</summary>
    public string? Description { get; set; }

    /// <summary>BT-129 Menge.</summary>
    public decimal Quantity { get; set; }

    /// <summary>BT-130 Einheit (UN/ECE Rec. 20), z. B. C62 = Stück, HUR = Stunde.</summary>
    public string UnitCode { get; set; } = "C62";

    /// <summary>BT-146 Einzelpreis netto.</summary>
    public decimal UnitPrice { get; set; }

    /// <summary>BT-151 Umsatzsteuerkategorie (UNTDID 5305), z. B. S = Normal-/ermäßigter Satz.</summary>
    public string VatCategory { get; set; } = "S";

    /// <summary>BT-152 Umsatzsteuersatz in Prozent.</summary>
    public decimal VatRate { get; set; }

    /// <summary>
    /// BT-131 Nettobetrag der Position as stated in a received invoice. <c>null</c> = compute
    /// as <c>round(Quantity × UnitPrice, 2)</c>.
    /// </summary>
    public decimal? StatedNetAmount { get; set; }
}

/// <summary>BG-22 Gesamtbeträge plus BG-23 Umsatzsteueraufschlüsselung.</summary>
public sealed class InvoiceTotals
{
    /// <summary>BT-106 Summe der Nettobeträge aller Positionen.</summary>
    public decimal LineTotal { get; set; }

    /// <summary>BT-107 Summe der Abschläge auf Dokumentenebene.</summary>
    public decimal AllowanceTotal { get; set; }

    /// <summary>BT-108 Summe der Zuschläge auf Dokumentenebene.</summary>
    public decimal ChargeTotal { get; set; }

    /// <summary>BT-109 Rechnungsbetrag ohne Umsatzsteuer.</summary>
    public decimal TaxBasisTotal { get; set; }

    /// <summary>BT-110 Umsatzsteuer gesamt.</summary>
    public decimal TaxTotal { get; set; }

    /// <summary>BT-112 Rechnungsbetrag inkl. Umsatzsteuer.</summary>
    public decimal GrandTotal { get; set; }

    /// <summary>BT-113 Vorauszahlungen.</summary>
    public decimal PrepaidAmount { get; set; }

    /// <summary>BT-114 Rundungsbetrag.</summary>
    public decimal RoundingAmount { get; set; }

    /// <summary>BT-115 Fälliger Zahlungsbetrag.</summary>
    public decimal DuePayable { get; set; }

    /// <summary>BG-23 Aufschlüsselung je Umsatzsteuerkategorie und -satz.</summary>
    public List<VatBreakdown> VatBreakdown { get; set; } = new();
}

/// <summary>BG-23 Umsatzsteueraufschlüsselung.</summary>
public sealed class VatBreakdown
{
    /// <summary>BT-118 Kategorie.</summary>
    public string Category { get; set; } = "S";

    /// <summary>BT-119 Satz in Prozent.</summary>
    public decimal Rate { get; set; }

    /// <summary>BT-116 Steuerbemessungsgrundlage.</summary>
    public decimal TaxableAmount { get; set; }

    /// <summary>BT-117 Steuerbetrag.</summary>
    public decimal TaxAmount { get; set; }

    /// <summary>BT-120 Befreiungsgrund.</summary>
    public string? ExemptionReason { get; set; }

    /// <summary>BT-121 Code für den Befreiungsgrund.</summary>
    public string? ExemptionReasonCode { get; set; }
}
