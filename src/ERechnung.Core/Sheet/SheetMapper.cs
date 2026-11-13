using System.Globalization;
using ERechnung.Core.Codes;
using ERechnung.Core.Model;
using ERechnung.Core.Validation;

namespace ERechnung.Core.Sheet;

/// <summary>Result of reading an invoice sheet: the invoice plus problems with unreadable cell values.</summary>
public sealed record SheetReadResult(Invoice Invoice, IReadOnlyList<ValidationMessage> Problems);

/// <summary>
/// Converts raw Excel cell values (as delivered by <c>Range.Value2</c>: string, double, bool or null) to the
/// <see cref="Invoice"/> model and back. Kept free of any Excel dependency so it can be unit tested.
/// </summary>
public static class SheetMapper
{
    private static readonly CultureInfo De = CultureInfo.GetCultureInfo("de-DE");

    public static SheetReadResult Read(
        IReadOnlyDictionary<string, object?> fields,
        IEnumerable<IReadOnlyDictionary<string, object?>> lineRows)
    {
        var problems = new List<ValidationMessage>();
        var r = new Reader(fields, problems);

        var inv = new Invoice
        {
            Number = r.Text(FieldKeys.Number),
            IssueDate = r.Date(FieldKeys.IssueDate),
            TypeCode = r.Code(FieldKeys.TypeCode) ?? "380",
            CurrencyCode = r.Code(FieldKeys.Currency) ?? "EUR",
            DeliveryDate = r.Date(FieldKeys.DeliveryDate),
            PeriodStart = r.Date(FieldKeys.PeriodStart),
            PeriodEnd = r.Date(FieldKeys.PeriodEnd),
            DueDate = r.Date(FieldKeys.DueDate),
            PaymentTerms = r.Text(FieldKeys.PaymentTerms),
            BuyerReference = r.Text(FieldKeys.BuyerReference),
            OrderReference = r.Text(FieldKeys.OrderReference),
            ContractReference = r.Text(FieldKeys.ContractReference),
            PrecedingInvoiceNumber = r.Text(FieldKeys.PrecedingInvoiceNumber),
            PrecedingInvoiceDate = r.Date(FieldKeys.PrecedingInvoiceDate),
            VatExemptionReason = r.Text(FieldKeys.VatExemptionReason),
            VatExemptionReasonCode = r.Code(FieldKeys.VatExemptionReasonCode),
            PrepaidAmount = r.Decimal(FieldKeys.PrepaidAmount) ?? 0m,
            Note = r.Text(FieldKeys.Note),
            Seller = new Party
            {
                Name = r.Text(FieldKeys.SellerName),
                Identifier = r.Text(FieldKeys.SellerIdentifier),
                Street = r.Text(FieldKeys.SellerStreet),
                PostCode = r.Text(FieldKeys.SellerPostCode),
                City = r.Text(FieldKeys.SellerCity),
                CountryCode = r.Code(FieldKeys.SellerCountry) ?? "DE",
                VatId = r.Text(FieldKeys.SellerVatId)?.Replace(" ", ""),
                TaxNumber = r.Text(FieldKeys.SellerTaxNumber),
                ElectronicAddress = r.Text(FieldKeys.SellerElectronicAddress),
                ElectronicAddressScheme = r.Code(FieldKeys.SellerElectronicAddressScheme),
                ContactName = r.Text(FieldKeys.SellerContactName),
                ContactPhone = r.Text(FieldKeys.SellerContactPhone),
                ContactEmail = r.Text(FieldKeys.SellerContactEmail),
            },
            Buyer = new Party
            {
                Name = r.Text(FieldKeys.BuyerName),
                Identifier = r.Text(FieldKeys.BuyerIdentifier),
                Street = r.Text(FieldKeys.BuyerStreet),
                PostCode = r.Text(FieldKeys.BuyerPostCode),
                City = r.Text(FieldKeys.BuyerCity),
                CountryCode = r.Code(FieldKeys.BuyerCountry) ?? "DE",
                VatId = r.Text(FieldKeys.BuyerVatId)?.Replace(" ", ""),
                ElectronicAddress = r.Text(FieldKeys.BuyerElectronicAddress),
                ElectronicAddressScheme = r.Code(FieldKeys.BuyerElectronicAddressScheme),
                ContactName = r.Text(FieldKeys.BuyerContactName),
                ContactPhone = r.Text(FieldKeys.BuyerContactPhone),
                ContactEmail = r.Text(FieldKeys.BuyerContactEmail),
            },
            Payment = new PaymentInstructions
            {
                MeansCode = r.Code(FieldKeys.PaymentMeans) ?? "58",
                Iban = r.Text(FieldKeys.PaymentIban),
                Bic = r.Text(FieldKeys.PaymentBic),
                AccountName = r.Text(FieldKeys.PaymentAccountName),
                RemittanceInformation = r.Text(FieldKeys.PaymentRemittance),
                MandateReference = r.Text(FieldKeys.PaymentMandate),
                CreditorId = r.Text(FieldKeys.PaymentCreditorId),
                DebitedIban = r.Text(FieldKeys.PaymentDebitedIban),
            },
        };

        var rowNo = 0;
        foreach (var row in lineRows)
        {
            rowNo++;
            // The template pre-fills Pos., Einheit, USt-Kategorie and USt-Satz; a row only counts once it has content.
            if (ContentColumns.All(k => IsBlank(row.GetValueOrDefault(k)))) continue;

            var lr = new Reader(row, problems, $"Position (Zeile {rowNo})");
            inv.Lines.Add(new InvoiceLine
            {
                Id = lr.Text(LineKeys.Id) ?? rowNo.ToString(CultureInfo.InvariantCulture),
                ArticleNumber = lr.Text(LineKeys.ArticleNumber),
                Name = lr.Text(LineKeys.Name),
                Description = lr.Text(LineKeys.Description),
                Quantity = lr.Decimal(LineKeys.Quantity) ?? 0m,
                UnitCode = lr.Code(LineKeys.Unit) ?? "C62",
                UnitPrice = lr.Decimal(LineKeys.UnitPrice) ?? 0m,
                VatCategory = lr.Code(LineKeys.VatCategory) ?? "S",
                VatRate = lr.Decimal(LineKeys.VatRate) ?? 0m,
            });
        }

        return new SheetReadResult(inv, problems);
    }

    /// <summary>
    /// Converts an invoice to cell values: dates as <see cref="DateTime"/>, amounts as <see cref="decimal"/>,
    /// codes as display text ("380 - Rechnung"). Missing values are <c>null</c>.
    /// </summary>
    public static (Dictionary<string, object?> Fields, List<Dictionary<string, object?>> Lines) Write(Invoice inv)
    {
        var f = new Dictionary<string, object?>
        {
            [FieldKeys.Number] = inv.Number,
            [FieldKeys.IssueDate] = inv.IssueDate,
            [FieldKeys.TypeCode] = CodeLists.Display(CodeLists.InvoiceTypes, inv.TypeCode),
            [FieldKeys.Currency] = CodeLists.Display(CodeLists.Currencies, inv.CurrencyCode),
            [FieldKeys.DeliveryDate] = inv.DeliveryDate,
            [FieldKeys.PeriodStart] = inv.PeriodStart,
            [FieldKeys.PeriodEnd] = inv.PeriodEnd,
            [FieldKeys.DueDate] = inv.DueDate,
            [FieldKeys.PaymentTerms] = inv.PaymentTerms,
            [FieldKeys.BuyerReference] = inv.BuyerReference,
            [FieldKeys.OrderReference] = inv.OrderReference,
            [FieldKeys.ContractReference] = inv.ContractReference,
            [FieldKeys.PrecedingInvoiceNumber] = inv.PrecedingInvoiceNumber,
            [FieldKeys.PrecedingInvoiceDate] = inv.PrecedingInvoiceDate,
            [FieldKeys.VatExemptionReason] = inv.VatExemptionReason,
            [FieldKeys.VatExemptionReasonCode] = inv.VatExemptionReasonCode,
            [FieldKeys.PrepaidAmount] = inv.PrepaidAmount == 0 ? null : inv.PrepaidAmount,
            [FieldKeys.Note] = inv.Note,

            [FieldKeys.SellerName] = inv.Seller.Name,
            [FieldKeys.SellerStreet] = inv.Seller.Street,
            [FieldKeys.SellerPostCode] = inv.Seller.PostCode,
            [FieldKeys.SellerCity] = inv.Seller.City,
            [FieldKeys.SellerCountry] = CodeLists.Display(CodeLists.Countries, inv.Seller.CountryCode),
            [FieldKeys.SellerVatId] = inv.Seller.VatId,
            [FieldKeys.SellerTaxNumber] = inv.Seller.TaxNumber,
            [FieldKeys.SellerIdentifier] = inv.Seller.Identifier,
            [FieldKeys.SellerElectronicAddress] = inv.Seller.ElectronicAddress,
            [FieldKeys.SellerElectronicAddressScheme] = CodeLists.Display(CodeLists.ElectronicAddressSchemes, inv.Seller.ElectronicAddressScheme),
            [FieldKeys.SellerContactName] = inv.Seller.ContactName,
            [FieldKeys.SellerContactPhone] = inv.Seller.ContactPhone,
            [FieldKeys.SellerContactEmail] = inv.Seller.ContactEmail,

            [FieldKeys.BuyerName] = inv.Buyer.Name,
            [FieldKeys.BuyerIdentifier] = inv.Buyer.Identifier,
            [FieldKeys.BuyerStreet] = inv.Buyer.Street,
            [FieldKeys.BuyerPostCode] = inv.Buyer.PostCode,
            [FieldKeys.BuyerCity] = inv.Buyer.City,
            [FieldKeys.BuyerCountry] = CodeLists.Display(CodeLists.Countries, inv.Buyer.CountryCode),
            [FieldKeys.BuyerVatId] = inv.Buyer.VatId,
            [FieldKeys.BuyerElectronicAddress] = inv.Buyer.ElectronicAddress,
            [FieldKeys.BuyerElectronicAddressScheme] = CodeLists.Display(CodeLists.ElectronicAddressSchemes, inv.Buyer.ElectronicAddressScheme),
            [FieldKeys.BuyerContactName] = inv.Buyer.ContactName,
            [FieldKeys.BuyerContactPhone] = inv.Buyer.ContactPhone,
            [FieldKeys.BuyerContactEmail] = inv.Buyer.ContactEmail,

            [FieldKeys.PaymentMeans] = CodeLists.Display(CodeLists.PaymentMeans, inv.Payment.MeansCode),
            [FieldKeys.PaymentIban] = inv.Payment.Iban,
            [FieldKeys.PaymentBic] = inv.Payment.Bic,
            [FieldKeys.PaymentAccountName] = inv.Payment.AccountName,
            [FieldKeys.PaymentRemittance] = inv.Payment.RemittanceInformation,
            [FieldKeys.PaymentMandate] = inv.Payment.MandateReference,
            [FieldKeys.PaymentCreditorId] = inv.Payment.CreditorId,
            [FieldKeys.PaymentDebitedIban] = inv.Payment.DebitedIban,
        };

        var lines = inv.Lines.Select(l => new Dictionary<string, object?>
        {
            [LineKeys.Id] = l.Id,
            [LineKeys.ArticleNumber] = l.ArticleNumber,
            [LineKeys.Name] = l.Name,
            [LineKeys.Description] = l.Description,
            [LineKeys.Quantity] = l.Quantity,
            [LineKeys.Unit] = CodeLists.Display(CodeLists.Units, l.UnitCode),
            [LineKeys.UnitPrice] = l.UnitPrice,
            [LineKeys.VatCategory] = CodeLists.Display(CodeLists.VatCategories, l.VatCategory),
            [LineKeys.VatRate] = l.VatRate,
            [LineKeys.NetAmount] = Calculation.InvoiceCalculator.LineNet(l),
        }).ToList();

        return (f, lines);
    }

    private static readonly string[] ContentColumns =
        { LineKeys.Name, LineKeys.Description, LineKeys.ArticleNumber, LineKeys.Quantity, LineKeys.UnitPrice };

    private static bool IsBlank(object? v) => v is null || (v is string s && string.IsNullOrWhiteSpace(s));

    /// <summary>Converts single cell values and records problems with a German message.</summary>
    private sealed class Reader
    {
        private readonly IReadOnlyDictionary<string, object?> _values;
        private readonly List<ValidationMessage> _problems;
        private readonly string? _context;

        public Reader(IReadOnlyDictionary<string, object?> values, List<ValidationMessage> problems, string? context = null)
        {
            _values = values;
            _problems = problems;
            _context = context;
        }

        private object? Raw(string key) => _values.TryGetValue(key, out var v) ? v : null;

        public string? Text(string key)
        {
            var v = Raw(key);
            return v switch
            {
                null => null,
                string s => string.IsNullOrWhiteSpace(s) ? null : s.Trim(),
                double d => d.ToString("0.############", CultureInfo.InvariantCulture),   // numbers typed into text fields (PLZ, Kundennummer)
                decimal m => m.ToString(CultureInfo.InvariantCulture),
                DateTime dt => dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                _ => v.ToString()?.Trim(),
            };
        }

        public string? Code(string key) => CodeLists.ParseCode(Text(key));

        public DateTime? Date(string key)
        {
            var v = Raw(key);
            switch (v)
            {
                case null:
                    return null;
                case DateTime dt:
                    return dt.Date;
                case double d:
                    try { return DateTime.FromOADate(d).Date; }
                    catch (ArgumentException) { break; }
                case string s when string.IsNullOrWhiteSpace(s):
                    return null;
                case string s:
                    var formats = new[] { "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "yyyy-MM-dd" };
                    if (DateTime.TryParseExact(s.Trim(), formats, De, DateTimeStyles.None, out var parsed)) return parsed.Date;
                    break;
            }
            Problem(key, $"\"{v}\" ist kein gültiges Datum (erwartet z. B. 31.12.2027).");
            return null;
        }

        public decimal? Decimal(string key)
        {
            var v = Raw(key);
            switch (v)
            {
                case null:
                    return null;
                case double d:
                    return (decimal)d;
                case decimal m:
                    return m;
                case int i:
                    return i;
                case string s when string.IsNullOrWhiteSpace(s):
                    return null;
                case string s:
                    var t = s.Trim().Replace("€", "").Replace("%", "").Trim();
                    if (decimal.TryParse(t, NumberStyles.Number, De, out var de)) return de;
                    if (decimal.TryParse(t, NumberStyles.Number, CultureInfo.InvariantCulture, out var inv)) return inv;
                    break;
            }
            Problem(key, $"\"{v}\" ist keine gültige Zahl.");
            return null;
        }

        private void Problem(string key, string text)
        {
            var label = SheetLayout.Field(key)?.Label ?? SheetLayout.LineColumns.FirstOrDefault(c => c.Key == key)?.Header ?? key;
            var prefix = _context is null ? label : $"{_context}, {label}";
            _problems.Add(new ValidationMessage(Severity.Error, "EINGABE", $"{prefix}: {text}", MessageSource.BuiltIn,
                _context is null ? key : FieldKeys.Lines));
        }
    }
}
