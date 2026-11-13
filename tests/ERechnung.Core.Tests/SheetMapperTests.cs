using ERechnung.Core.Samples;
using ERechnung.Core.Sheet;
using ERechnung.Core.Validation;

namespace ERechnung.Core.Tests;

public class SheetMapperTests
{
    [Fact]
    public void Read_ConvertsExcelValueTypes()
    {
        var fields = new Dictionary<string, object?>
        {
            [FieldKeys.Number] = "R-1",
            [FieldKeys.IssueDate] = new DateTime(2027, 1, 31).ToOADate(),     // date cell → double
            [FieldKeys.DueDate] = "28.02.2027",                                // date typed as text
            [FieldKeys.TypeCode] = "381 - Gutschrift (Rechnungsminderung)",    // drop-down display text
            [FieldKeys.Currency] = "eur",
            [FieldKeys.SellerPostCode] = 1067d,                                // number typed into text field
            [FieldKeys.SellerVatId] = "DE 123 456 789",
            [FieldKeys.PrepaidAmount] = "1.234,50",                            // German number text
            [FieldKeys.PaymentMeans] = "59 - SEPA-Lastschrift",
        };
        var rows = new[]
        {
            new Dictionary<string, object?>
            {
                [LineKeys.Name] = "Beratung",
                [LineKeys.Quantity] = 2.5d,
                [LineKeys.Unit] = "HUR - Stunde",
                [LineKeys.UnitPrice] = "100,10 €",
                [LineKeys.VatCategory] = "S - Normal-/ermäßigter Satz",
                [LineKeys.VatRate] = 19d,
                [LineKeys.NetAmount] = 999d,                                   // computed column: ignored
            },
            new Dictionary<string, object?>                                    // template defaults only: skipped
            {
                [LineKeys.Id] = "2", [LineKeys.Unit] = "C62 - Stück", [LineKeys.VatCategory] = "S - Normal-/ermäßigter Satz",
                [LineKeys.VatRate] = 19d, [LineKeys.NetAmount] = 0d,
            },
        };

        var (inv, problems) = SheetMapper.Read(fields, rows);

        Assert.Empty(problems);
        Assert.Equal(new DateTime(2027, 1, 31), inv.IssueDate);
        Assert.Equal(new DateTime(2027, 2, 28), inv.DueDate);
        Assert.Equal("381", inv.TypeCode);
        Assert.Equal("EUR", inv.CurrencyCode);
        Assert.Equal("1067", inv.Seller.PostCode);
        Assert.Equal("DE123456789", inv.Seller.VatId);
        Assert.Equal(1234.50m, inv.PrepaidAmount);
        Assert.Equal("59", inv.Payment.MeansCode);
        var line = Assert.Single(inv.Lines);
        Assert.Equal("1", line.Id);
        Assert.Equal(2.5m, line.Quantity);
        Assert.Equal("HUR", line.UnitCode);
        Assert.Equal(100.10m, line.UnitPrice);
        Assert.Null(line.StatedNetAmount);
    }

    [Fact]
    public void Read_ReportsUnparseableValues_WithFieldKey()
    {
        var fields = new Dictionary<string, object?> { [FieldKeys.IssueDate] = "gestern", [FieldKeys.PrepaidAmount] = "viel" };

        var (_, problems) = SheetMapper.Read(fields, Array.Empty<IReadOnlyDictionary<string, object?>>());

        Assert.Contains(problems, p => p.FieldKey == FieldKeys.IssueDate && p.Text.Contains("kein gültiges Datum"));
        Assert.Contains(problems, p => p.FieldKey == FieldKeys.PrepaidAmount && p.Text.Contains("keine gültige Zahl"));
    }

    [Fact]
    public void WriteThenRead_RoundTrips()
    {
        var original = DemoInvoices.Standard();
        var (fields, lines) = SheetMapper.Write(original);
        // Excel hands dates and numbers back as doubles
        var asExcel = fields.ToDictionary(kv => kv.Key, kv => kv.Value switch
        {
            DateTime d => (object?)d.ToOADate(),
            decimal m => (double)m,
            var v => v,
        });
        var rows = lines.Select(l => (IReadOnlyDictionary<string, object?>)l.ToDictionary(kv => kv.Key,
            kv => kv.Value is decimal m ? (object?)(double)m : kv.Value)).ToList();

        var (inv, problems) = SheetMapper.Read(asExcel, rows);

        Assert.Empty(problems);
        Assert.Equal(original.Number, inv.Number);
        Assert.Equal(original.IssueDate, inv.IssueDate);
        Assert.Equal(original.Seller.Name, inv.Seller.Name);
        Assert.Equal(original.Payment.Iban, inv.Payment.Iban);
        Assert.Equal(original.Lines.Count, inv.Lines.Count);
        Assert.Equal(original.Lines[0].UnitPrice, inv.Lines[0].UnitPrice);
        Assert.Equal(0, BuiltInValidator.Validate(inv, ValidationProfile.XRechnung).ErrorCount);
    }

    [Fact]
    public void Layout_KeysAndNamesAreUnique()
    {
        var keys = SheetLayout.AllFields.Select(f => f.Key).ToList();
        Assert.Equal(keys.Count, keys.Distinct().Count());
        var names = keys.Select(SheetLayout.NameFor).ToList();
        Assert.All(names, n => Assert.Matches("^ER_[A-Za-z_]+$", n));
        Assert.Equal(SheetLayout.LineColumns.Count, SheetLayout.LineColumns.Select(c => c.Header).Distinct().Count());
    }
}
