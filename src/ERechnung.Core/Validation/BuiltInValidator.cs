using System.Numerics;
using System.Text.RegularExpressions;
using ERechnung.Core.Calculation;
using ERechnung.Core.Codes;
using ERechnung.Core.Model;
using ERechnung.Core.Sheet;

namespace ERechnung.Core.Validation;

/// <summary>Which rule set applies.</summary>
public enum ValidationProfile
{
    /// <summary>EN 16931 core rules plus the German CIUS XRechnung 3.0 (BR-DE-*).</summary>
    XRechnung,

    /// <summary>EN 16931 core rules only (e.g. ZUGFeRD profile EN 16931 for B2B).</summary>
    EN16931,
}

/// <summary>
/// Fast, offline checks of the most important EN 16931 and XRechnung rules, with German messages that point to
/// the sheet fields. This is a pre-check for the user, not a replacement for the official KoSIT validator:
/// it covers the rules typical spreadsheet input gets wrong, not the full Schematron.
/// Rule ids follow EN 16931 (BR-*, BR-CO-*, BR-S-* …) and XRechnung (BR-DE-*) where a rule maps 1:1.
/// </summary>
public static class BuiltInValidator
{
    private static readonly HashSet<string> XRechnungTypeCodes = new() { "326", "380", "384", "389", "381", "875", "876", "877" };
    private static readonly HashSet<string> CreditTransferCodes = new() { "30", "58" };

    public static ValidationReport Validate(Invoice inv, ValidationProfile profile, ValidationReport? into = null)
    {
        var r = into ?? new ValidationReport();
        var xr = profile == ValidationProfile.XRechnung;

        // --- Document level --------------------------------------------------------
        Required(r, inv.Number, "BR-02", "Die Rechnungsnummer fehlt.", FieldKeys.Number);
        if (inv.IssueDate is null) r.Add(Severity.Error, "BR-03", "Das Rechnungsdatum fehlt.", FieldKeys.IssueDate);
        Required(r, inv.TypeCode, "BR-04", "Die Rechnungsart fehlt.", FieldKeys.TypeCode);
        Required(r, inv.CurrencyCode, "BR-05", "Die Währung fehlt.", FieldKeys.Currency);
        if (xr && !string.IsNullOrWhiteSpace(inv.TypeCode) && !XRechnungTypeCodes.Contains(inv.TypeCode))
            r.Add(Severity.Error, "BR-DE-17", $"Rechnungsart {inv.TypeCode} ist in der XRechnung nicht erlaubt (erlaubt: {string.Join(", ", XRechnungTypeCodes)}).", FieldKeys.TypeCode);

        if (xr) Required(r, inv.BuyerReference, "BR-DE-15",
            "Die Käuferreferenz (BT-10) fehlt. Bei Behörden ist das die Leitweg-ID, sonst eine vom Käufer vorgegebene Referenz.", FieldKeys.BuyerReference);
        if (inv.BuyerReference is { } br && br.Contains('-') && char.IsDigit(br.Trim()[0]) && !ElectronicAddressSchemes.IsLeitwegId(br))
            r.Add(Severity.Warning, "LEITWEG-ID", $"\"{br}\" sieht aus wie eine Leitweg-ID, hat aber nicht das Format Grobadressierung-Feinadressierung-Prüfziffer (z. B. 04011000-12345-34).", FieldKeys.BuyerReference);

        if (inv.TypeCode is "384" or "381" && string.IsNullOrWhiteSpace(inv.PrecedingInvoiceNumber))
            r.Add(Severity.Warning, inv.TypeCode == "384" ? "BR-DE-26" : "BT-25",
                "Bei Korrekturen und Gutschriften sollte die ursprüngliche Rechnungsnummer angegeben werden.", FieldKeys.PrecedingInvoiceNumber);

        if (inv.DueDate is { } due && inv.IssueDate is { } issued && due < issued)
            r.Add(Severity.Warning, "DATUM", "Das Fälligkeitsdatum liegt vor dem Rechnungsdatum.", FieldKeys.DueDate);
        if (inv.PeriodStart is { } ps && inv.PeriodEnd is { } pe && pe < ps)
            r.Add(Severity.Error, "BR-29", "Das Ende des Leistungszeitraums liegt vor dessen Beginn.", FieldKeys.PeriodEnd);
        if (inv.DeliveryDate is null && inv.PeriodStart is null && inv.PeriodEnd is null)
            r.Add(Severity.Warning, "UStG-14", "Kein Leistungsdatum bzw. Leistungszeitraum angegeben (§14 Abs. 4 Nr. 6 UStG verlangt den Zeitpunkt der Leistung).", FieldKeys.DeliveryDate);

        // --- Seller --------------------------------------------------------------------
        var s = inv.Seller;
        Required(r, s.Name, "BR-06", "Der Name des Verkäufers fehlt.", FieldKeys.SellerName);
        Required(r, s.CountryCode, "BR-09", "Das Land des Verkäufers fehlt.", FieldKeys.SellerCountry);
        if (xr)
        {
            Required(r, s.City, "BR-DE-3", "Der Ort des Verkäufers fehlt.", FieldKeys.SellerCity);
            Required(r, s.PostCode, "BR-DE-4", "Die PLZ des Verkäufers fehlt.", FieldKeys.SellerPostCode);
            Required(r, s.ContactName, "BR-DE-5", "Der Ansprechpartner des Verkäufers fehlt.", FieldKeys.SellerContactName);
            Required(r, s.ContactPhone, "BR-DE-6", "Die Telefonnummer des Verkäufers fehlt.", FieldKeys.SellerContactPhone);
            Required(r, s.ContactEmail, "BR-DE-7", "Die E-Mail-Adresse des Verkäufers fehlt.", FieldKeys.SellerContactEmail);
            Required(r, s.ElectronicAddress, "PEPPOL-EN16931-R020", "Die elektronische Adresse des Verkäufers (BT-34) fehlt – in XRechnung 3.0 Pflicht. Meist die E-Mail-Adresse für Rechnungen.", FieldKeys.SellerElectronicAddress);
        }
        if (s.ContactPhone is { } phone && Regex.Matches(phone, "[0-9]").Count < 3)
            r.Add(Severity.Warning, "BR-DE-27", "Die Telefonnummer des Verkäufers sollte mindestens drei Ziffern enthalten.", FieldKeys.SellerContactPhone);
        if (s.ContactEmail is { } mail && !IsEmail(mail))
            r.Add(Severity.Warning, "BR-DE-28", $"\"{mail}\" ist keine gültige E-Mail-Adresse.", FieldKeys.SellerContactEmail);
        if (string.IsNullOrWhiteSpace(s.VatId) && string.IsNullOrWhiteSpace(s.Identifier))
            r.Add(Severity.Error, "BR-CO-26", "Ohne USt-IdNr. muss eine Verkäufer-Kennung (BT-29) angegeben werden, z. B. die Lieferantennummer beim Kunden. Die Steuernummer allein genügt nicht.", FieldKeys.SellerIdentifier);
        CheckVatId(r, s.VatId, FieldKeys.SellerVatId, "Verkäufers");
        CheckElectronicAddress(r, s, FieldKeys.SellerElectronicAddress, "Verkäufers");

        // --- Buyer ---------------------------------------------------------------------
        var b = inv.Buyer;
        Required(r, b.Name, "BR-07", "Der Name des Käufers fehlt.", FieldKeys.BuyerName);
        Required(r, b.CountryCode, "BR-11", "Das Land des Käufers fehlt.", FieldKeys.BuyerCountry);
        if (xr)
        {
            Required(r, b.City, "BR-DE-8", "Der Ort des Käufers fehlt.", FieldKeys.BuyerCity);
            Required(r, b.PostCode, "BR-DE-9", "Die PLZ des Käufers fehlt.", FieldKeys.BuyerPostCode);
            Required(r, b.ElectronicAddress, "PEPPOL-EN16931-R010", "Die elektronische Adresse des Käufers (BT-49) fehlt – in XRechnung 3.0 Pflicht (E-Mail oder Leitweg-ID).", FieldKeys.BuyerElectronicAddress);
        }
        CheckVatId(r, b.VatId, FieldKeys.BuyerVatId, "Käufers");
        CheckElectronicAddress(r, b, FieldKeys.BuyerElectronicAddress, "Käufers");

        // --- Lines -----------------------------------------------------------------------
        if (inv.Lines.Count == 0)
            r.Add(Severity.Error, "BR-16", "Die Rechnung enthält keine Positionen.", FieldKeys.Lines);

        var ids = new HashSet<string>();
        foreach (var l in inv.Lines)
        {
            var where = $"Position {l.Id}";
            if (!string.IsNullOrWhiteSpace(l.Id) && !ids.Add(l.Id))
                r.Add(Severity.Warning, "BR-21", $"{where}: Die Positionsnummer kommt mehrfach vor.", FieldKeys.Lines, where);
            if (string.IsNullOrWhiteSpace(l.Name))
                r.Add(Severity.Error, "BR-25", $"{where}: Die Bezeichnung fehlt.", FieldKeys.Lines, where);
            if (l.Quantity == 0)
                r.Add(Severity.Warning, "BR-22", $"{where}: Die Menge ist 0.", FieldKeys.Lines, where);
            if (string.IsNullOrWhiteSpace(l.UnitCode))
                r.Add(Severity.Error, "BR-23", $"{where}: Die Einheit fehlt.", FieldKeys.Lines, where);
            if (l.UnitPrice < 0)
                r.Add(Severity.Error, "BR-27", $"{where}: Der Einzelpreis darf nicht negativ sein. Für Gutschriften eine negative Menge oder Rechnungsart 381 verwenden.", FieldKeys.Lines, where);
            CheckLineVat(r, l, where);
        }

        // --- VAT -------------------------------------------------------------------------
        var cats = inv.Lines.Select(l => (l.VatCategory ?? "").ToUpperInvariant()).ToHashSet();
        var sellerHasTaxId = !string.IsNullOrWhiteSpace(s.VatId) || !string.IsNullOrWhiteSpace(s.TaxNumber);
        if (cats.Overlaps(new[] { "S", "Z", "E", "AE", "K", "G" }) && !sellerHasTaxId)
            r.Add(Severity.Error, xr ? "BR-DE-16" : "BR-S-02", "USt-IdNr. oder Steuernummer des Verkäufers fehlt.", FieldKeys.SellerVatId);
        if (cats.Contains("AE"))
        {
            if (string.IsNullOrWhiteSpace(b.VatId))
                r.Add(Severity.Error, "BR-AE-02", "Reverse Charge (AE) erfordert die USt-IdNr. des Käufers.", FieldKeys.BuyerVatId);
            if (string.IsNullOrWhiteSpace(s.VatId))
                r.Add(Severity.Error, "BR-AE-02", "Reverse Charge (AE) erfordert die USt-IdNr. des Verkäufers (eine Steuernummer reicht nicht).", FieldKeys.SellerVatId);
        }
        if (cats.Contains("K"))
        {
            if (string.IsNullOrWhiteSpace(b.VatId))
                r.Add(Severity.Error, "BR-IC-02", "Innergemeinschaftliche Lieferung (K) erfordert die USt-IdNr. des Käufers.", FieldKeys.BuyerVatId);
            if (string.IsNullOrWhiteSpace(s.VatId))
                r.Add(Severity.Error, "BR-IC-02", "Innergemeinschaftliche Lieferung (K) erfordert die USt-IdNr. des Verkäufers.", FieldKeys.SellerVatId);
        }
        if (cats.Overlaps(CodeLists.CategoriesRequiringExemptionReason) &&
            string.IsNullOrWhiteSpace(inv.VatExemptionReason) && string.IsNullOrWhiteSpace(inv.VatExemptionReasonCode))
        {
            var which = string.Join(", ", cats.Intersect(CodeLists.CategoriesRequiringExemptionReason));
            r.Add(Severity.Error, "BR-E-10", $"Für die USt-Kategorie {which} muss ein Grund der Steuerbefreiung angegeben werden (z. B. \"Kleinunternehmer gemäß §19 UStG\" oder \"Steuerschuldnerschaft des Leistungsempfängers\").", FieldKeys.VatExemptionReason);
        }
        if (!string.IsNullOrWhiteSpace(inv.VatExemptionReasonCode) &&
            !Regex.IsMatch(inv.VatExemptionReasonCode, "^VATEX-[A-Z]{2}-[A-Z0-9-]+$", RegexOptions.IgnoreCase))
            r.Add(Severity.Error, "BR-CL-22", $"\"{inv.VatExemptionReasonCode}\" ist kein Code der VATEX-Liste (z. B. VATEX-EU-AE, VATEX-EU-IC).", FieldKeys.VatExemptionReasonCode);

        // --- Payment ---------------------------------------------------------------------
        var p = inv.Payment;
        if (xr) Required(r, p.MeansCode, "BR-DE-1", "Die Zahlungsart fehlt.", FieldKeys.PaymentMeans);
        if (CreditTransferCodes.Contains(p.MeansCode ?? ""))
        {
            if (string.IsNullOrWhiteSpace(p.Iban))
                r.Add(xr ? Severity.Error : Severity.Warning, "BR-DE-23-a", "Bei Zahlungsart Überweisung muss die IBAN (BT-84) angegeben werden.", FieldKeys.PaymentIban);
            else if (p.MeansCode == "58" && !IsValidIban(p.Iban))
                r.Add(Severity.Warning, "BR-DE-19", $"Die IBAN \"{p.Iban}\" ist ungültig (Prüfziffer oder Format) – eine Zahlung würde fehlschlagen.", FieldKeys.PaymentIban);
        }
        if (p.MeansCode is "48" or "54" or "55")
            r.Add(Severity.Error, "BR-DE-24-a", "Kartenzahlung erfordert Kartendaten (BG-18), die diese Vorlage nicht erfasst. Bitte eine andere Zahlungsart wählen.", FieldKeys.PaymentMeans);
        if (p.MeansCode == "59")
        {
            Required(r, p.MandateReference, "PEPPOL-EN16931-R061", "Bei SEPA-Lastschrift fehlt die Mandatsreferenz (BT-89).", FieldKeys.PaymentMandate);
            Required(r, p.CreditorId, "BR-DE-30", "Bei SEPA-Lastschrift fehlt die Gläubiger-ID (BT-90).", FieldKeys.PaymentCreditorId);
            Required(r, p.DebitedIban, "BR-DE-31", "Bei SEPA-Lastschrift fehlt die IBAN des Zahlers (BT-91).", FieldKeys.PaymentDebitedIban);
            if (!string.IsNullOrWhiteSpace(p.DebitedIban) && !IsValidIban(p.DebitedIban))
                r.Add(Severity.Warning, "BR-DE-20", $"Die IBAN des Zahlers \"{p.DebitedIban}\" ist ungültig.", FieldKeys.PaymentDebitedIban);
        }

        // --- Totals --------------------------------------------------------------------
        var computed = InvoiceCalculator.Calculate(inv);
        if (computed.DuePayable > 0 && inv.DueDate is null && string.IsNullOrWhiteSpace(inv.PaymentTerms))
            r.Add(Severity.Error, "BR-CO-25", "Bei einem positiven Zahlbetrag muss ein Fälligkeitsdatum oder eine Zahlungsbedingung angegeben werden.", FieldKeys.DueDate);
        if (inv.StatedTotals is { } st) CheckStatedTotals(r, inv, st);

        return r;
    }

    // --- helpers ------------------------------------------------------------------------

    private static void Required(ValidationReport r, string? value, string code, string text, string fieldKey)
    {
        if (string.IsNullOrWhiteSpace(value)) r.Add(Severity.Error, code, text, fieldKey);
    }

    private static void CheckLineVat(ValidationReport r, InvoiceLine l, string where)
    {
        var cat = (l.VatCategory ?? "").ToUpperInvariant();
        if (cat.Length == 0)
        {
            r.Add(Severity.Error, "BR-CO-04", $"{where}: Die USt-Kategorie fehlt.", FieldKeys.Lines, where);
            return;
        }
        if (CodeLists.LabelOf(CodeLists.VatCategories, cat) is null)
        {
            r.Add(Severity.Error, "BR-CL-18", $"{where}: Unbekannte USt-Kategorie \"{cat}\".", FieldKeys.Lines, where);
            return;
        }
        switch (cat)
        {
            case "S" when l.VatRate <= 0:
                r.Add(Severity.Error, "BR-S-05", $"{where}: Kategorie S (Normalsatz) braucht einen Steuersatz > 0 %. Für 0 % die Kategorie Z oder E wählen.", FieldKeys.Lines, where);
                break;
            case "S" when l.VatRate is not (19m or 7m):
                r.Add(Severity.Warning, "USt-SATZ", $"{where}: Ungewöhnlicher Steuersatz {l.VatRate} % (in Deutschland üblich: 19 % oder 7 %).", FieldKeys.Lines, where);
                break;
            case "Z" or "E" or "AE" or "K" or "G" when l.VatRate != 0:
                r.Add(Severity.Error, $"BR-{(cat == "K" ? "IC" : cat)}-05", $"{where}: Bei Kategorie {cat} muss der Steuersatz 0 % sein.", FieldKeys.Lines, where);
                break;
            case "O" when l.VatRate != 0:
                r.Add(Severity.Error, "BR-O-05", $"{where}: Bei Kategorie O (nicht steuerbar) darf kein Steuersatz angegeben werden.", FieldKeys.Lines, where);
                break;
        }
    }

    private static void CheckVatId(ValidationReport r, string? vatId, string fieldKey, string whose)
    {
        if (string.IsNullOrWhiteSpace(vatId)) return;
        if (!Regex.IsMatch(vatId, "^[A-Z]{2}[0-9A-Z+*.]{2,13}$"))
            r.Add(Severity.Error, "BR-CO-09", $"Die USt-IdNr. des {whose} muss mit dem Ländercode beginnen (z. B. DE123456789).", fieldKey);
        else if (vatId.StartsWith("DE") && !Regex.IsMatch(vatId, "^DE[0-9]{9}$"))
            r.Add(Severity.Warning, "USt-IdNr", $"Eine deutsche USt-IdNr. hat das Format DE + 9 Ziffern (Eingabe: {vatId}).", fieldKey);
    }

    private static void CheckElectronicAddress(ValidationReport r, Party p, string fieldKey, string whose)
    {
        if (string.IsNullOrWhiteSpace(p.ElectronicAddress)) return;
        var scheme = CodeLists.ParseCode(p.ElectronicAddressScheme) ?? ElectronicAddressSchemes.Detect(p.ElectronicAddress);
        if (scheme == "EM" && !IsEmail(p.ElectronicAddress))
            r.Add(Severity.Error, "BT-34-1", $"Die elektronische Adresse des {whose} ist als E-Mail eingestuft, aber keine gültige E-Mail-Adresse.", fieldKey);
        if (scheme == "0204" && !ElectronicAddressSchemes.IsLeitwegId(p.ElectronicAddress))
            r.Add(Severity.Error, "BT-34-1", $"Die elektronische Adresse des {whose} ist keine gültige Leitweg-ID.", fieldKey);
    }

    private static void CheckStatedTotals(ValidationReport r, Invoice inv, InvoiceTotals st)
    {
        var lineSum = inv.Lines.Sum(InvoiceCalculator.LineNet);
        if (lineSum != st.LineTotal)
            r.Add(Severity.Error, "BR-CO-10", $"Summe der Positionen ({lineSum:N2}) ≠ angegebene Nettosumme ({st.LineTotal:N2}).", FieldKeys.Lines);
        var basis = st.LineTotal - st.AllowanceTotal + st.ChargeTotal;
        if (basis != st.TaxBasisTotal)
            r.Add(Severity.Error, "BR-CO-13", $"Rechnungsbetrag ohne USt ({st.TaxBasisTotal:N2}) ≠ Nettosumme − Abschläge + Zuschläge ({basis:N2}).");
        var vatSum = st.VatBreakdown.Sum(v => v.TaxAmount);
        if (vatSum != st.TaxTotal)
            r.Add(Severity.Error, "BR-CO-14", $"USt-Gesamtbetrag ({st.TaxTotal:N2}) ≠ Summe der USt-Aufschlüsselung ({vatSum:N2}).");
        if (st.TaxBasisTotal + st.TaxTotal != st.GrandTotal)
            r.Add(Severity.Error, "BR-CO-15", $"Bruttobetrag ({st.GrandTotal:N2}) ≠ Netto + USt ({st.TaxBasisTotal + st.TaxTotal:N2}).");
        if (st.GrandTotal - st.PrepaidAmount + st.RoundingAmount != st.DuePayable)
            r.Add(Severity.Error, "BR-CO-16", $"Zahlbetrag ({st.DuePayable:N2}) ≠ Brutto − Vorauszahlung + Rundung ({st.GrandTotal - st.PrepaidAmount + st.RoundingAmount:N2}).");
        foreach (var v in st.VatBreakdown.Where(v => v.Category == "S"))
        {
            var expected = InvoiceCalculator.Round2(v.TaxableAmount * v.Rate / 100m);
            if (Math.Abs(expected - v.TaxAmount) > 0.01m)
                r.Add(Severity.Error, "BR-CO-17", $"USt für {v.Rate} %: angegeben {v.TaxAmount:N2}, erwartet {expected:N2}.");
        }
    }

    internal static bool IsEmail(string s) => Regex.IsMatch(s.Trim(), @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

    /// <summary>ISO 13616 check: length 15–34, country code, mod-97 = 1.</summary>
    internal static bool IsValidIban(string? iban)
    {
        if (string.IsNullOrWhiteSpace(iban)) return false;
        var s = new string(iban.Where(c => !char.IsWhiteSpace(c)).ToArray()).ToUpperInvariant();
        if (!Regex.IsMatch(s, "^[A-Z]{2}[0-9]{2}[A-Z0-9]{11,30}$")) return false;
        if (s.StartsWith("DE") && s.Length != 22) return false;
        var rearranged = s[4..] + s[..4];
        var digits = string.Concat(rearranged.Select(c => char.IsLetter(c) ? (c - 'A' + 10).ToString() : c.ToString()));
        return BigInteger.Parse(digits) % 97 == 1;
    }
}
