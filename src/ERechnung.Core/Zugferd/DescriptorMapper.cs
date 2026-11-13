using ERechnung.Core.Calculation;
using ERechnung.Core.Codes;
using ERechnung.Core.Model;
using s2industries.ZUGFeRD;

namespace ERechnung.Core.Zugferd;

/// <summary>
/// Maps between the neutral <see cref="Invoice"/> model and ZUGFeRD-csharp's <see cref="InvoiceDescriptor"/>.
/// The descriptor is what the library serialises to XRechnung (UBL/CII) and ZUGFeRD/Factur-X.
/// </summary>
internal static class DescriptorMapper
{
    /// <summary>BT-23 Geschäftsprozess: Peppol BIS Billing, the value XRechnung expects.</summary>
    public const string BusinessProcessPeppolBilling = "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0";

    public static InvoiceDescriptor ToDescriptor(Invoice inv)
    {
        if (string.IsNullOrWhiteSpace(inv.Number))
            throw new InvoiceDataException("Rechnungsnummer (BT-1) fehlt.");
        if (inv.IssueDate is null)
            throw new InvoiceDataException("Rechnungsdatum (BT-2) fehlt.");

        var currency = EnumCodes.FromCode<CurrencyCodes>(inv.CurrencyCode, "Währung");
        var d = InvoiceDescriptor.CreateInvoice(inv.Number.Trim(), inv.IssueDate.Value.Date, currency);
        d.Type = EnumCodes.FromCode<InvoiceType>(inv.TypeCode, "Rechnungsart");
        // BT-23: required by the XRechnung CII rules (PEPPOL-EN16931-R001); UBL gets it as cbc:ProfileID.
        d.BusinessProcess = BusinessProcessPeppolBilling;

        d.ReferenceOrderNo = Clean(inv.BuyerReference);                     // BT-10
        if (Clean(inv.OrderReference) is { } order) d.SetBuyerOrderReferenceDocument(order);
        if (Clean(inv.ContractReference) is { } contract) d.SetContractReferencedDocument(contract, null);
        if (Clean(inv.PrecedingInvoiceNumber) is { } preceding)
            d.AddInvoiceReferencedDocument(preceding, inv.PrecedingInvoiceDate);
        if (Clean(inv.Note) is { } note) d.AddNote(note);

        // --- Parties --------------------------------------------------------
        var s = inv.Seller;
        d.SetSeller(Clean(s.Name), Clean(s.PostCode), Clean(s.City), Clean(s.Street),
            Country(s.CountryCode, "Land des Verkäufers"), id: Clean(s.Identifier));
        if (HasContact(s)) d.SetSellerContact(Clean(s.ContactName), null, Clean(s.ContactEmail), Clean(s.ContactPhone));
        if (Clean(s.VatId) is { } sVat) d.AddSellerTaxRegistration(sVat, TaxRegistrationSchemeID.VA);
        if (Clean(s.TaxNumber) is { } sTax) d.AddSellerTaxRegistration(sTax, TaxRegistrationSchemeID.FC);
        if (Clean(s.ElectronicAddress) is { } sAddr)
            d.SetSellerElectronicAddress(sAddr, Scheme(s, "Verkäufer"));

        var b = inv.Buyer;
        d.SetBuyer(Clean(b.Name), Clean(b.PostCode), Clean(b.City), Clean(b.Street),
            Country(b.CountryCode, "Land des Käufers"), id: Clean(b.Identifier));
        if (HasContact(b)) d.SetBuyerContact(Clean(b.ContactName), null, Clean(b.ContactEmail), Clean(b.ContactPhone));
        if (Clean(b.VatId) is { } bVat) d.AddBuyerTaxRegistration(bVat, TaxRegistrationSchemeID.VA);
        if (Clean(b.ElectronicAddress) is { } bAddr)
            d.SetBuyerElectronicAddress(bAddr, Scheme(b, "Käufer"));

        // --- Delivery ---------------------------------------------------------
        d.ActualDeliveryDate = inv.DeliveryDate?.Date;
        if (inv.PeriodStart is not null || inv.PeriodEnd is not null)
            d.SetBillingPeriod(inv.PeriodStart?.Date, inv.PeriodEnd?.Date);

        // --- Payment ----------------------------------------------------------
        var p = inv.Payment;
        var means = EnumCodes.FromCode<PaymentMeansTypeCodes>(p.MeansCode, "Zahlungsart");
        if (means == PaymentMeansTypeCodes.SEPADirectDebit)
        {
            d.SetPaymentMeansSepaDirectDebit(Clean(p.CreditorId), Clean(p.MandateReference));
            if (Clean(p.DebitedIban) is { } debited) d.AddDebitorFinancialAccount(NormalizeIban(debited), null);
        }
        else
        {
            d.SetPaymentMeans(means);
        }
        if (Clean(p.Iban) is { } iban && means != PaymentMeansTypeCodes.SEPADirectDebit)
            d.AddCreditorFinancialAccount(NormalizeIban(iban), Clean(p.Bic), name: Clean(p.AccountName));
        d.PaymentReference = Clean(p.RemittanceInformation);

        // With only a due date, CII would get an empty <ram:Description/>, which violates PEPPOL-EN16931-R008
        // ("no empty elements"), so a text is derived from the date.
        var termsText = Clean(inv.PaymentTerms) ?? (inv.DueDate is { } dueDate ? $"Zahlbar bis {dueDate:dd.MM.yyyy}." : null);
        if (termsText is not null)
            d.AddTradePaymentTerms(termsText, inv.DueDate?.Date);

        // --- Lines ------------------------------------------------------------
        var lineNo = 0;
        foreach (var l in inv.Lines)
        {
            lineNo++;
            var id = Clean(l.Id) ?? lineNo.ToString();
            var unit = EnumCodes.FromCode<QuantityCodes>(l.UnitCode, $"Position {id}: Einheit");
            var cat = EnumCodes.FromCode<TaxCategoryCodes>(l.VatCategory, $"Position {id}: USt-Kategorie");
            d.AddTradeLineItem(
                lineID: id,
                name: Clean(l.Name) ?? "",
                netUnitPrice: l.UnitPrice,
                unitCode: unit,
                description: Clean(l.Description),
                billedQuantity: l.Quantity,
                lineTotalAmount: InvoiceCalculator.LineNet(l),
                taxType: TaxTypes.VAT,
                categoryCode: cat,
                taxPercent: l.VatRate,
                sellerAssignedID: Clean(l.ArticleNumber));
        }

        // --- Totals (always computed, never taken from the sheet) --------------
        var t = InvoiceCalculator.Calculate(inv);
        foreach (var v in t.VatBreakdown)
        {
            TaxExemptionReasonCodes? reasonCode = null;
            if (v.ExemptionReasonCode is { } rc)
            {
                if (!EnumCodes.TryFromCode<TaxExemptionReasonCodes>(rc, out var parsed))
                    throw new InvoiceDataException($"Befreiungsgrund-Code \"{rc}\" ist kein gültiger VATEX-Code.");
                reasonCode = parsed;
            }
            d.AddApplicableTradeTax(v.TaxableAmount, v.Rate, v.TaxAmount, TaxTypes.VAT,
                EnumCodes.FromCode<TaxCategoryCodes>(v.Category, "USt-Kategorie"),
                exemptionReasonCode: reasonCode, exemptionReason: v.ExemptionReason);
        }
        d.SetTotals(
            lineTotalAmount: t.LineTotal,
            chargeTotalAmount: null,
            allowanceTotalAmount: null,
            taxBasisAmount: t.TaxBasisTotal,
            taxTotalAmount: t.TaxTotal,
            grandTotalAmount: t.GrandTotal,
            totalPrepaidAmount: t.PrepaidAmount != 0 ? t.PrepaidAmount : null,
            duePayableAmount: t.DuePayable);

        return d;
    }

    public static Invoice FromDescriptor(InvoiceDescriptor d)
    {
        var inv = new Invoice
        {
            Number = d.InvoiceNo,
            IssueDate = d.InvoiceDate,
            TypeCode = EnumCodes.ToCode(d.Type),
            CurrencyCode = EnumCodes.ToCode(d.Currency),
            BuyerReference = d.ReferenceOrderNo,
            OrderReference = d.OrderNo,
            ContractReference = d.ContractReferencedDocument?.ID,
            DeliveryDate = d.ActualDeliveryDate,
            PeriodStart = d.BillingPeriodStart,
            PeriodEnd = d.BillingPeriodEnd,
            Note = JoinNonEmpty(d.Notes?.Select(n => n.Content)),
        };

        var preceding = d.GetInvoiceReferencedDocuments()?.FirstOrDefault();
        if (preceding is not null)
        {
            inv.PrecedingInvoiceNumber = preceding.ID;
            inv.PrecedingInvoiceDate = preceding.IssueDateTime;
        }

        var terms = d.GetTradePaymentTerms() ?? new List<PaymentTerms>();
        inv.PaymentTerms = JoinNonEmpty(terms.Select(t => t.Description));
        inv.DueDate = terms.Select(t => t.DueDate).FirstOrDefault(x => x is not null);

        inv.Seller = FromParty(d.Seller, d.SellerContact, d.SellerElectronicAddress, d.SellerTaxRegistration);
        inv.Buyer = FromParty(d.Buyer, d.BuyerContact, d.BuyerElectronicAddress, d.BuyerTaxRegistration);

        // Payment
        var pay = new PaymentInstructions
        {
            MeansCode = d.PaymentMeans?.TypeCode is { } tc ? EnumCodes.ToCode(tc) : "",
            RemittanceInformation = d.PaymentReference,
            MandateReference = d.PaymentMeans?.SEPAMandateReference,
            CreditorId = d.PaymentMeans?.SEPACreditorIdentifier,
        };
        var creditor = d.CreditorBankAccounts?.FirstOrDefault();
        if (creditor is not null)
        {
            pay.Iban = creditor.IBAN;
            pay.Bic = creditor.BIC;
            pay.AccountName = creditor.Name;
        }
        pay.DebitedIban = d.DebitorBankAccounts?.FirstOrDefault()?.IBAN;
        inv.Payment = pay;

        // Lines
        foreach (var li in d.TradeLineItems ?? new List<TradeLineItem>())
        {
            inv.Lines.Add(new InvoiceLine
            {
                Id = li.AssociatedDocument?.LineID,
                ArticleNumber = li.SellerAssignedID,
                Name = li.Name,
                Description = li.Description,
                Quantity = li.BilledQuantity,
                UnitCode = li.UnitCode is { } u ? EnumCodes.ToCode(u) : "",
                UnitPrice = li.NetUnitPrice,
                VatCategory = li.TaxCategoryCode is { } c ? EnumCodes.ToCode(c) : "",
                VatRate = li.TaxPercent,
                StatedNetAmount = li.LineTotalAmount,
            });
        }

        // Totals + VAT breakdown as stated in the document
        var totals = new InvoiceTotals
        {
            LineTotal = d.LineTotalAmount ?? 0m,
            AllowanceTotal = d.AllowanceTotalAmount ?? 0m,
            ChargeTotal = d.ChargeTotalAmount ?? 0m,
            TaxBasisTotal = d.TaxBasisAmount ?? 0m,
            TaxTotal = d.TaxTotalAmount ?? 0m,
            GrandTotal = d.GrandTotalAmount ?? 0m,
            PrepaidAmount = d.TotalPrepaidAmount ?? 0m,
            RoundingAmount = d.RoundingAmount ?? 0m,
            DuePayable = d.DuePayableAmount ?? 0m,
        };
        foreach (var tax in d.Taxes ?? new List<Tax>())
        {
            totals.VatBreakdown.Add(new VatBreakdown
            {
                Category = tax.CategoryCode is { } c ? EnumCodes.ToCode(c) : "",
                Rate = tax.Percent,
                TaxableAmount = tax.BasisAmount,
                TaxAmount = tax.TaxAmount,
                ExemptionReason = tax.ExemptionReason,
                ExemptionReasonCode = tax.ExemptionReasonCode is { } rc ? EnumCodes.ToCode(rc) : null,
            });
        }
        inv.StatedTotals = totals;
        inv.PrepaidAmount = totals.PrepaidAmount;

        var exempt = totals.VatBreakdown.FirstOrDefault(v => v.ExemptionReason is not null || v.ExemptionReasonCode is not null);
        inv.VatExemptionReason = exempt?.ExemptionReason;
        inv.VatExemptionReasonCode = exempt?.ExemptionReasonCode;
        return inv;
    }

    private static Model.Party FromParty(s2industries.ZUGFeRD.Party? p, Contact? c, ElectronicAddress? addr, List<TaxRegistration>? regs)
    {
        var party = new Model.Party
        {
            Name = p?.Name,
            Identifier = p?.ID?.ID,
            Street = p?.Street,
            PostCode = p?.Postcode,
            City = p?.City,
            CountryCode = p?.Country is { } cc ? EnumCodes.ToCode(cc) : "",
            ContactName = c?.Name,
            ContactPhone = c?.PhoneNo,
            ContactEmail = c?.EmailAddress,
            ElectronicAddress = addr?.Address,
            ElectronicAddressScheme = addr is null ? null : EnumCodes.ToCode(addr.ElectronicAddressSchemeID),
        };
        foreach (var r in regs ?? new List<TaxRegistration>())
        {
            if (r.SchemeID == TaxRegistrationSchemeID.VA) party.VatId ??= r.No;
            else if (r.SchemeID == TaxRegistrationSchemeID.FC) party.TaxNumber ??= r.No;
        }
        return party;
    }

    private static string? Clean(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private static bool HasContact(Model.Party p) =>
        Clean(p.ContactName) is not null || Clean(p.ContactEmail) is not null || Clean(p.ContactPhone) is not null;

    private static CountryCodes Country(string? code, string label) =>
        EnumCodes.FromCode<CountryCodes>(string.IsNullOrWhiteSpace(code) ? "DE" : code, label);

    private static ElectronicAddressSchemeIdentifiers Scheme(Model.Party p, string who)
    {
        var code = Clean(p.ElectronicAddressScheme) ?? ElectronicAddressSchemes.Detect(p.ElectronicAddress);
        return EnumCodes.FromCode<ElectronicAddressSchemeIdentifiers>(code, $"Schema der elektronischen Adresse ({who})");
    }

    internal static string NormalizeIban(string iban) => new string(iban.Where(ch => !char.IsWhiteSpace(ch)).ToArray()).ToUpperInvariant();

    private static string? JoinNonEmpty(IEnumerable<string?>? parts)
    {
        if (parts is null) return null;
        var list = parts.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x!.Trim()).ToList();
        return list.Count == 0 ? null : string.Join(Environment.NewLine, list);
    }
}
