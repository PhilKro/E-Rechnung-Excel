using ERechnung.Core.Codes;

namespace ERechnung.Core.Sheet;

/// <summary>Stable keys for every header field. Excel names are derived from them (<see cref="SheetLayout.NameFor"/>).</summary>
public static class FieldKeys
{
    public const string Number = "Number";
    public const string IssueDate = "IssueDate";
    public const string TypeCode = "TypeCode";
    public const string Currency = "Currency";
    public const string DeliveryDate = "DeliveryDate";
    public const string PeriodStart = "PeriodStart";
    public const string PeriodEnd = "PeriodEnd";
    public const string DueDate = "DueDate";
    public const string BuyerReference = "BuyerReference";
    public const string OrderReference = "OrderReference";
    public const string ContractReference = "ContractReference";
    public const string PrecedingInvoiceNumber = "PrecedingInvoiceNumber";
    public const string PrecedingInvoiceDate = "PrecedingInvoiceDate";
    public const string PaymentTerms = "PaymentTerms";
    public const string Note = "Note";
    public const string VatExemptionReason = "VatExemptionReason";
    public const string VatExemptionReasonCode = "VatExemptionReasonCode";
    public const string PrepaidAmount = "PrepaidAmount";

    public const string SellerName = "Seller.Name";
    public const string SellerIdentifier = "Seller.Identifier";
    public const string SellerStreet = "Seller.Street";
    public const string SellerPostCode = "Seller.PostCode";
    public const string SellerCity = "Seller.City";
    public const string SellerCountry = "Seller.Country";
    public const string SellerVatId = "Seller.VatId";
    public const string SellerTaxNumber = "Seller.TaxNumber";
    public const string SellerElectronicAddress = "Seller.ElectronicAddress";
    public const string SellerElectronicAddressScheme = "Seller.ElectronicAddressScheme";
    public const string SellerContactName = "Seller.ContactName";
    public const string SellerContactPhone = "Seller.ContactPhone";
    public const string SellerContactEmail = "Seller.ContactEmail";

    public const string BuyerName = "Buyer.Name";
    public const string BuyerIdentifier = "Buyer.Identifier";
    public const string BuyerStreet = "Buyer.Street";
    public const string BuyerPostCode = "Buyer.PostCode";
    public const string BuyerCity = "Buyer.City";
    public const string BuyerCountry = "Buyer.Country";
    public const string BuyerVatId = "Buyer.VatId";
    public const string BuyerElectronicAddress = "Buyer.ElectronicAddress";
    public const string BuyerElectronicAddressScheme = "Buyer.ElectronicAddressScheme";
    public const string BuyerContactName = "Buyer.ContactName";
    public const string BuyerContactPhone = "Buyer.ContactPhone";
    public const string BuyerContactEmail = "Buyer.ContactEmail";

    public const string PaymentMeans = "Payment.MeansCode";
    public const string PaymentIban = "Payment.Iban";
    public const string PaymentBic = "Payment.Bic";
    public const string PaymentAccountName = "Payment.AccountName";
    public const string PaymentRemittance = "Payment.Remittance";
    public const string PaymentMandate = "Payment.MandateReference";
    public const string PaymentCreditorId = "Payment.CreditorId";
    public const string PaymentDebitedIban = "Payment.DebitedIban";

    /// <summary>Pseudo key for messages about the line table as a whole.</summary>
    public const string Lines = "Lines";
}

/// <summary>Keys of the line table columns.</summary>
public static class LineKeys
{
    public const string Id = "Id";
    public const string ArticleNumber = "ArticleNumber";
    public const string Name = "Name";
    public const string Description = "Description";
    public const string Quantity = "Quantity";
    public const string Unit = "Unit";
    public const string UnitPrice = "UnitPrice";
    public const string VatCategory = "VatCategory";
    public const string VatRate = "VatRate";
    public const string NetAmount = "NetAmount";
}

public enum FieldKind
{
    Text,
    Date,
    Decimal,
    Code,
}

/// <param name="Required">Mandatory for XRechnung; shown bold in the sheet.</param>
public sealed record FieldDef(
    string Key,
    string Label,
    string BusinessTerm,
    FieldKind Kind = FieldKind.Text,
    IReadOnlyList<CodeEntry>? Codes = null,
    bool Required = false,
    string? Hint = null);

public sealed record SectionDef(string Title, IReadOnlyList<FieldDef> Fields);

/// <param name="Computed">Computed by a formula in the sheet (line net amount); ignored when reading.</param>
public sealed record ColumnDef(
    string Key,
    string Header,
    string BusinessTerm,
    FieldKind Kind = FieldKind.Text,
    IReadOnlyList<CodeEntry>? Codes = null,
    bool Computed = false);

/// <summary>
/// Describes the invoice worksheet: which fields exist, in which order, with which labels and drop-downs.
/// The Excel add-in renders this layout; <see cref="SheetMapper"/> converts cell values to/from <see cref="Model.Invoice"/>.
/// Changing a label is safe; changing a key breaks existing workbooks (names are derived from keys).
/// </summary>
public static class SheetLayout
{
    /// <summary>Prefix of every sheet-scoped Excel name the add-in creates.</summary>
    public const string NamePrefix = "ER_";

    /// <summary>Sheet-scoped name that marks a worksheet as an invoice sheet.</summary>
    public const string MarkerName = NamePrefix + "Rechnungsblatt";

    /// <summary>Prefix of the line table (ListObject) name; a counter is appended to keep it workbook-unique.</summary>
    public const string LineTablePrefix = "ER_Positionen";

    /// <summary>Excel name for a field key, e.g. "Seller.Name" → "ER_Seller_Name".</summary>
    public static string NameFor(string key) => NamePrefix + key.Replace('.', '_');

    public static readonly IReadOnlyList<SectionDef> Sections = new[]
    {
        new SectionDef("Rechnung", new[]
        {
            new FieldDef(FieldKeys.Number, "Rechnungsnummer", "BT-1", Required: true),
            new FieldDef(FieldKeys.IssueDate, "Rechnungsdatum", "BT-2", FieldKind.Date, Required: true),
            new FieldDef(FieldKeys.TypeCode, "Rechnungsart", "BT-3", FieldKind.Code, CodeLists.InvoiceTypes, Required: true),
            new FieldDef(FieldKeys.Currency, "Währung", "BT-5", FieldKind.Code, CodeLists.Currencies, Required: true),
            new FieldDef(FieldKeys.DeliveryDate, "Leistungsdatum", "BT-72", FieldKind.Date, Hint: "oder Leistungszeitraum"),
            new FieldDef(FieldKeys.PeriodStart, "Leistungszeitraum von", "BT-73", FieldKind.Date),
            new FieldDef(FieldKeys.PeriodEnd, "Leistungszeitraum bis", "BT-74", FieldKind.Date),
            new FieldDef(FieldKeys.DueDate, "Fälligkeitsdatum", "BT-9", FieldKind.Date),
            new FieldDef(FieldKeys.PaymentTerms, "Zahlungsbedingungen", "BT-20", Hint: "z. B. \"Zahlbar innerhalb 14 Tagen ohne Abzug\""),
            new FieldDef(FieldKeys.BuyerReference, "Käuferreferenz / Leitweg-ID", "BT-10", Required: true, Hint: "Pflicht bei XRechnung; Behörden: Leitweg-ID"),
            new FieldDef(FieldKeys.OrderReference, "Bestellnummer", "BT-13"),
            new FieldDef(FieldKeys.ContractReference, "Vertragsnummer", "BT-12"),
            new FieldDef(FieldKeys.PrecedingInvoiceNumber, "Bezug: Rechnungsnummer", "BT-25", Hint: "bei Korrektur/Gutschrift"),
            new FieldDef(FieldKeys.PrecedingInvoiceDate, "Bezug: Rechnungsdatum", "BT-26", FieldKind.Date),
            new FieldDef(FieldKeys.VatExemptionReason, "Grund der Steuerbefreiung", "BT-120", Hint: "Pflicht bei USt-Kategorie E, AE, K, G, O"),
            new FieldDef(FieldKeys.VatExemptionReasonCode, "Befreiungsgrund-Code", "BT-121", Hint: "optional, z. B. VATEX-EU-AE"),
            new FieldDef(FieldKeys.PrepaidAmount, "Bereits gezahlt", "BT-113", FieldKind.Decimal),
            new FieldDef(FieldKeys.Note, "Bemerkung", "BT-22"),
        }),
        new SectionDef("Verkäufer (Rechnungssteller)", new[]
        {
            new FieldDef(FieldKeys.SellerName, "Name / Firma", "BT-27", Required: true),
            new FieldDef(FieldKeys.SellerStreet, "Straße, Hausnr.", "BT-35"),
            new FieldDef(FieldKeys.SellerPostCode, "PLZ", "BT-38", Required: true),
            new FieldDef(FieldKeys.SellerCity, "Ort", "BT-37", Required: true),
            new FieldDef(FieldKeys.SellerCountry, "Land", "BT-40", FieldKind.Code, CodeLists.Countries, Required: true),
            new FieldDef(FieldKeys.SellerVatId, "USt-IdNr.", "BT-31", Hint: "USt-IdNr. oder Steuernummer"),
            new FieldDef(FieldKeys.SellerTaxNumber, "Steuernummer", "BT-32"),
            new FieldDef(FieldKeys.SellerIdentifier, "Verkäufer-Kennung", "BT-29", Hint: "Pflicht ohne USt-IdNr., z. B. Lieferantennr. beim Kunden"),
            new FieldDef(FieldKeys.SellerElectronicAddress, "Elektronische Adresse", "BT-34", Required: true, Hint: "meist E-Mail für Rechnungen"),
            new FieldDef(FieldKeys.SellerElectronicAddressScheme, "Schema elektr. Adresse", "BT-34-1", FieldKind.Code, CodeLists.ElectronicAddressSchemes, Hint: "leer = automatisch"),
            new FieldDef(FieldKeys.SellerContactName, "Ansprechpartner", "BT-41", Required: true),
            new FieldDef(FieldKeys.SellerContactPhone, "Telefon", "BT-42", Required: true),
            new FieldDef(FieldKeys.SellerContactEmail, "E-Mail", "BT-43", Required: true),
        }),
        new SectionDef("Käufer (Rechnungsempfänger)", new[]
        {
            new FieldDef(FieldKeys.BuyerName, "Name / Firma", "BT-44", Required: true),
            new FieldDef(FieldKeys.BuyerIdentifier, "Kundennummer", "BT-46"),
            new FieldDef(FieldKeys.BuyerStreet, "Straße, Hausnr.", "BT-50"),
            new FieldDef(FieldKeys.BuyerPostCode, "PLZ", "BT-53", Required: true),
            new FieldDef(FieldKeys.BuyerCity, "Ort", "BT-52", Required: true),
            new FieldDef(FieldKeys.BuyerCountry, "Land", "BT-55", FieldKind.Code, CodeLists.Countries, Required: true),
            new FieldDef(FieldKeys.BuyerVatId, "USt-IdNr.", "BT-48", Hint: "Pflicht bei Reverse Charge / innergem. Lieferung"),
            new FieldDef(FieldKeys.BuyerElectronicAddress, "Elektronische Adresse", "BT-49", Required: true, Hint: "E-Mail oder Leitweg-ID"),
            new FieldDef(FieldKeys.BuyerElectronicAddressScheme, "Schema elektr. Adresse", "BT-49-1", FieldKind.Code, CodeLists.ElectronicAddressSchemes, Hint: "leer = automatisch"),
            new FieldDef(FieldKeys.BuyerContactName, "Ansprechpartner", "BT-56"),
            new FieldDef(FieldKeys.BuyerContactPhone, "Telefon", "BT-57"),
            new FieldDef(FieldKeys.BuyerContactEmail, "E-Mail", "BT-58"),
        }),
        new SectionDef("Zahlung", new[]
        {
            new FieldDef(FieldKeys.PaymentMeans, "Zahlungsart", "BT-81", FieldKind.Code, CodeLists.PaymentMeans, Required: true),
            new FieldDef(FieldKeys.PaymentIban, "IBAN", "BT-84", Hint: "Pflicht bei Überweisung"),
            new FieldDef(FieldKeys.PaymentBic, "BIC", "BT-86"),
            new FieldDef(FieldKeys.PaymentAccountName, "Kontoinhaber", "BT-85"),
            new FieldDef(FieldKeys.PaymentRemittance, "Verwendungszweck", "BT-83"),
            new FieldDef(FieldKeys.PaymentMandate, "Mandatsreferenz", "BT-89", Hint: "nur bei Lastschrift"),
            new FieldDef(FieldKeys.PaymentCreditorId, "Gläubiger-ID", "BT-90", Hint: "nur bei Lastschrift"),
            new FieldDef(FieldKeys.PaymentDebitedIban, "IBAN des Zahlers", "BT-91", Hint: "nur bei Lastschrift"),
        }),
    };

    /// <summary>
    /// Line table columns in sheet order. The table sits below the header fields and shares columns A–J with them,
    /// so the order is chosen to fit those widths (wide text columns under label/value/hint).
    /// Columns are matched by header text when reading, so reordering is safe; renaming headers breaks old sheets.
    /// </summary>
    public static readonly IReadOnlyList<ColumnDef> LineColumns = new[]
    {
        new ColumnDef(LineKeys.Id, "Pos.", "BT-126"),
        new ColumnDef(LineKeys.Name, "Bezeichnung", "BT-153"),
        new ColumnDef(LineKeys.Description, "Beschreibung", "BT-154"),
        new ColumnDef(LineKeys.ArticleNumber, "Artikelnr.", "BT-155"),
        new ColumnDef(LineKeys.Quantity, "Menge", "BT-129", FieldKind.Decimal),
        new ColumnDef(LineKeys.Unit, "Einheit", "BT-130", FieldKind.Code, CodeLists.Units),
        new ColumnDef(LineKeys.UnitPrice, "Einzelpreis netto", "BT-146", FieldKind.Decimal),
        new ColumnDef(LineKeys.VatCategory, "USt-Kategorie", "BT-151", FieldKind.Code, CodeLists.VatCategories),
        new ColumnDef(LineKeys.VatRate, "USt-Satz %", "BT-152", FieldKind.Decimal),
        new ColumnDef(LineKeys.NetAmount, "Nettobetrag", "BT-131", FieldKind.Decimal, Computed: true),
    };

    public static IEnumerable<FieldDef> AllFields => Sections.SelectMany(s => s.Fields);

    public static FieldDef? Field(string key) => AllFields.FirstOrDefault(f => f.Key == key);
}
