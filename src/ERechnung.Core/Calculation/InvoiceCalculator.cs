using ERechnung.Core.Codes;
using ERechnung.Core.Model;

namespace ERechnung.Core.Calculation;

/// <summary>
/// Computes line amounts, the VAT breakdown and the document totals the way EN 16931 expects:
/// <list type="bullet">
/// <item>Positionsbetrag (BT-131) = round(Menge × Einzelpreis, 2)</item>
/// <item>Steuer je Kategorie/Satz (BT-117) = round(Summe der Positionsbeträge × Satz / 100, 2) — rounded once per group, not per line</item>
/// <item>BT-109 = BT-106 − BT-107 + BT-108, BT-112 = BT-109 + BT-110, BT-115 = BT-112 − BT-113 + BT-114</item>
/// </list>
/// The Excel sheet shows line amounts via formulas, but these numbers are the authoritative ones written to the XML.
/// </summary>
public static class InvoiceCalculator
{
    /// <summary>Commercial rounding to cents.</summary>
    public static decimal Round2(decimal value) => Math.Round(value, 2, MidpointRounding.AwayFromZero);

    /// <summary>Line net amount: the stated amount for received invoices, otherwise quantity × price.</summary>
    public static decimal LineNet(InvoiceLine line) => line.StatedNetAmount ?? Round2(line.Quantity * line.UnitPrice);

    public static InvoiceTotals Calculate(Invoice invoice)
    {
        var totals = new InvoiceTotals();

        var groups = invoice.Lines
            .GroupBy(l => (Category: (l.VatCategory ?? "S").Trim().ToUpperInvariant(), l.VatRate))
            .OrderBy(g => g.Key.Category == "S" ? 0 : 1)
            .ThenByDescending(g => g.Key.VatRate);

        foreach (var g in groups)
        {
            var basis = g.Sum(LineNet);
            var vb = new VatBreakdown
            {
                Category = g.Key.Category,
                Rate = g.Key.VatRate,
                TaxableAmount = basis,
                TaxAmount = Round2(basis * g.Key.VatRate / 100m),
            };
            if (CodeLists.CategoriesRequiringExemptionReason.Contains(vb.Category))
            {
                vb.ExemptionReason = string.IsNullOrWhiteSpace(invoice.VatExemptionReason) ? null : invoice.VatExemptionReason.Trim();
                vb.ExemptionReasonCode = string.IsNullOrWhiteSpace(invoice.VatExemptionReasonCode) ? null : invoice.VatExemptionReasonCode.Trim();
            }
            totals.VatBreakdown.Add(vb);
        }

        totals.LineTotal = invoice.Lines.Sum(LineNet);
        totals.AllowanceTotal = 0m;
        totals.ChargeTotal = 0m;
        totals.TaxBasisTotal = totals.LineTotal - totals.AllowanceTotal + totals.ChargeTotal;
        totals.TaxTotal = totals.VatBreakdown.Sum(v => v.TaxAmount);
        totals.GrandTotal = totals.TaxBasisTotal + totals.TaxTotal;
        totals.PrepaidAmount = Round2(invoice.PrepaidAmount);
        totals.RoundingAmount = 0m;
        totals.DuePayable = totals.GrandTotal - totals.PrepaidAmount + totals.RoundingAmount;
        return totals;
    }
}
