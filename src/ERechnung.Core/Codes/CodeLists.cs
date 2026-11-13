using System.Text.RegularExpressions;

namespace ERechnung.Core.Codes;

/// <summary>A code with a German display label, used for Excel drop-downs and for display after import.</summary>
public sealed record CodeEntry(string Code, string Label)
{
    /// <summary>Text shown in Excel cells and drop-downs, e.g. "380 - Rechnung".</summary>
    public string Display => $"{Code} - {Label}";
}

/// <summary>
/// The code lists offered in the Excel template. They are deliberately short: only codes that make sense for
/// typical German B2B/B2G invoices. Received invoices may contain other valid codes; those are kept as-is.
/// </summary>
public static class CodeLists
{
    /// <summary>BT-3, restricted to the codes XRechnung allows (BR-DE-17).</summary>
    public static readonly IReadOnlyList<CodeEntry> InvoiceTypes = new[]
    {
        new CodeEntry("380", "Rechnung"),
        new CodeEntry("381", "Gutschrift (Rechnungsminderung)"),
        new CodeEntry("384", "Rechnungskorrektur"),
        new CodeEntry("326", "Teilrechnung"),
        new CodeEntry("389", "Gutschriftsverfahren (Selbstfakturierung)"),
        new CodeEntry("875", "Teilrechnung Bauleistung"),
        new CodeEntry("876", "Teilschlussrechnung Bauleistung"),
        new CodeEntry("877", "Schlussrechnung Bauleistung"),
    };

    /// <summary>BT-151 / BT-118 (UNTDID 5305, EN 16931 subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> VatCategories = new[]
    {
        new CodeEntry("S", "Normal-/ermäßigter Satz"),
        new CodeEntry("Z", "Nullsatz"),
        new CodeEntry("E", "Steuerbefreit (z. B. §19 UStG, §4 UStG)"),
        new CodeEntry("AE", "Reverse Charge (§13b UStG)"),
        new CodeEntry("K", "Innergemeinschaftliche Lieferung"),
        new CodeEntry("G", "Ausfuhr (Drittland)"),
        new CodeEntry("O", "Nicht steuerbar"),
    };

    /// <summary>Categories that require an exemption reason (BT-120 or BT-121).</summary>
    public static readonly IReadOnlySet<string> CategoriesRequiringExemptionReason =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "E", "AE", "K", "G", "O" };

    /// <summary>BT-81 (UNTDID 4461, common subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> PaymentMeans = new[]
    {
        new CodeEntry("58", "SEPA-Überweisung"),
        new CodeEntry("59", "SEPA-Lastschrift"),
        new CodeEntry("30", "Überweisung (nicht SEPA)"),
        new CodeEntry("10", "Barzahlung"),
        // 48/54/55 (Karte) deliberately missing: XRechnung then requires card data (BR-DE-24-a) the template does not collect.
        new CodeEntry("68", "Online-Zahlungsdienst"),
        new CodeEntry("1", "Nicht definiert"),
    };

    /// <summary>BT-130 (UN/ECE Recommendation 20, common subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> Units = new[]
    {
        new CodeEntry("C62", "Stück"),
        new CodeEntry("H87", "Stück (Teil)"),
        new CodeEntry("HUR", "Stunde"),
        new CodeEntry("MIN", "Minute"),
        new CodeEntry("DAY", "Tag"),
        new CodeEntry("WEE", "Woche"),
        new CodeEntry("MON", "Monat"),
        new CodeEntry("ANN", "Jahr"),
        new CodeEntry("LS", "Pauschale"),
        new CodeEntry("SET", "Satz"),
        new CodeEntry("KGM", "Kilogramm"),
        new CodeEntry("TNE", "Tonne"),
        new CodeEntry("MTR", "Meter"),
        new CodeEntry("KMT", "Kilometer"),
        new CodeEntry("MTK", "Quadratmeter"),
        new CodeEntry("MTQ", "Kubikmeter"),
        new CodeEntry("LTR", "Liter"),
        new CodeEntry("KWH", "Kilowattstunde"),
        new CodeEntry("XPK", "Packung"),
        new CodeEntry("P1", "Prozent"),
    };

    /// <summary>BT-5 (ISO 4217, subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> Currencies = new[]
    {
        new CodeEntry("EUR", "Euro"),
        new CodeEntry("CHF", "Schweizer Franken"),
        new CodeEntry("USD", "US-Dollar"),
        new CodeEntry("GBP", "Pfund Sterling"),
        new CodeEntry("PLN", "Złoty"),
        new CodeEntry("CZK", "Tschechische Krone"),
        new CodeEntry("DKK", "Dänische Krone"),
        new CodeEntry("SEK", "Schwedische Krone"),
    };

    /// <summary>BT-40 / BT-55 (ISO 3166-1, subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> Countries = new[]
    {
        new CodeEntry("DE", "Deutschland"),
        new CodeEntry("AT", "Österreich"),
        new CodeEntry("CH", "Schweiz"),
        new CodeEntry("FR", "Frankreich"),
        new CodeEntry("NL", "Niederlande"),
        new CodeEntry("BE", "Belgien"),
        new CodeEntry("LU", "Luxemburg"),
        new CodeEntry("IT", "Italien"),
        new CodeEntry("ES", "Spanien"),
        new CodeEntry("PL", "Polen"),
        new CodeEntry("CZ", "Tschechien"),
        new CodeEntry("DK", "Dänemark"),
        new CodeEntry("SE", "Schweden"),
        new CodeEntry("GB", "Vereinigtes Königreich"),
        new CodeEntry("US", "Vereinigte Staaten"),
    };

    /// <summary>Schemes for BT-34 / BT-49 (CEF EAS, subset).</summary>
    public static readonly IReadOnlyList<CodeEntry> ElectronicAddressSchemes = new[]
    {
        new CodeEntry("EM", "E-Mail"),
        new CodeEntry("0204", "Leitweg-ID"),
        new CodeEntry("9930", "USt-IdNr. Deutschland"),
        new CodeEntry("0088", "GLN"),
        new CodeEntry("0060", "DUNS"),
    };

    /// <summary>Returns the German label for a code, or <c>null</c> if the code is not in the list.</summary>
    public static string? LabelOf(IReadOnlyList<CodeEntry> list, string? code) =>
        list.FirstOrDefault(e => string.Equals(e.Code, code, StringComparison.OrdinalIgnoreCase))?.Label;

    /// <summary>"380" → "380 - Rechnung"; unknown codes are returned unchanged.</summary>
    public static string Display(IReadOnlyList<CodeEntry> list, string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return "";
        var label = LabelOf(list, code);
        return label is null ? code : $"{code} - {label}";
    }

    /// <summary>
    /// Extracts the code from a cell value: "380 - Rechnung" → "380", " s " → "S".
    /// Everything before the first " - " (or the first whitespace) is the code.
    /// </summary>
    public static string? ParseCode(string? cellText)
    {
        if (string.IsNullOrWhiteSpace(cellText)) return null;
        var t = cellText.Trim();
        var dash = t.IndexOf(" - ", StringComparison.Ordinal);
        if (dash > 0) t = t[..dash];
        var space = t.IndexOfAny(new[] { ' ', '\t' });
        if (space > 0) t = t[..space];
        return t.ToUpperInvariant();
    }
}

/// <summary>Helpers for electronic addresses (BT-34 / BT-49).</summary>
public static class ElectronicAddressSchemes
{
    private static readonly Regex LeitwegId = new(@"^[0-9]{2,12}(-[0-9A-Za-z]{1,30})?-[0-9]{2}$", RegexOptions.Compiled);

    /// <summary>Checks the formal structure of a Leitweg-ID (Grobadressierung[-Feinadressierung]-Prüfziffer).</summary>
    public static bool IsLeitwegId(string? value) => value is not null && LeitwegId.IsMatch(value.Trim());

    /// <summary>
    /// Picks a scheme when none was given: e-mail addresses → EM, Leitweg-IDs → 0204,
    /// German VAT ids → 9930. Falls back to EM.
    /// </summary>
    public static string Detect(string? address)
    {
        var a = address?.Trim() ?? "";
        if (a.Contains('@')) return "EM";
        if (IsLeitwegId(a)) return "0204";
        if (Regex.IsMatch(a, @"^DE[0-9]{9}$", RegexOptions.IgnoreCase)) return "9930";
        return "EM";
    }
}
