using ERechnung.Core.Calculation;
using ERechnung.Core.Codes;
using ERechnung.Core.IO;
using ERechnung.Core.Validation;
using Excel = Microsoft.Office.Interop.Excel;
using static ERechnung.ExcelAddIn.ExcelHelpers;

namespace ERechnung.ExcelAddIn;

/// <summary>
/// "Rechnungseingang": a log of received e-invoices, one table row per imported file, so incoming invoices can be
/// filtered, summed and booked in Excel.
/// </summary>
internal static class InboxSheet
{
    public const string SheetName = "Rechnungseingang";
    public const string TableName = "ER_Eingang";

    private static readonly (string Header, double Width, string Format)[] Columns =
    {
        ("Importiert am", 16, "dd.mm.yyyy hh:mm"),
        ("Datei", 30, "@"),
        ("Format", 26, "@"),
        ("Rechnungsnr.", 16, "@"),
        ("Rechnungsdatum", 13, "dd.mm.yyyy"),
        ("Art", 18, "@"),
        ("Verkäufer", 28, "@"),
        ("USt-IdNr. Verkäufer", 17, "@"),
        ("Käuferreferenz", 18, "@"),
        ("Netto", 13, "#,##0.00"),
        ("USt", 12, "#,##0.00"),
        ("Brutto", 13, "#,##0.00"),
        ("Zahlbetrag", 13, "#,##0.00"),
        ("Währung", 8, "@"),
        ("Fällig am", 12, "dd.mm.yyyy"),
        ("IBAN", 26, "@"),
        ("Verwendungszweck", 22, "@"),
        ("Prüfergebnis", 12, "@"),
        ("Fehler", 8, "0"),
        ("Warnungen", 10, "0"),
        ("Pfad", 50, "@"),
    };

    public static Excel.Worksheet EnsureSheet(Excel.Workbook wb)
    {
        var ws = FindSheet(wb, SheetName);
        if (ws is not null && ws.ListObjects.Cast<Excel.ListObject>().Any(l => l.Name == TableName)) return ws;

        ws ??= AddSheetAtEnd(wb, SheetName);
        Batch(() =>
        {
            ws.Cells.Font.Name = "Calibri";
            ws.Cells.Font.Size = 10;
            for (var i = 0; i < Columns.Length; i++)
            {
                Cell(ws, 1, i + 1).Value2 = Columns[i].Header;
                ((Excel.Range)ws.Columns[i + 1]).ColumnWidth = Columns[i].Width;
                Cell(ws, 2, i + 1).NumberFormat = Columns[i].Format;
            }
            var lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, Range(ws, 1, 1, 2, Columns.Length),
                Type.Missing, Excel.XlYesNoGuess.xlYes);
            lo.Name = TableName;
            lo.TableStyle = "TableStyleMedium9";
        });
        return ws;
    }

    public static void Append(Excel.Workbook wb, string path, ImportedInvoice imported, ValidationReport report)
    {
        var ws = EnsureSheet(wb);
        var lo = ws.ListObjects.Cast<Excel.ListObject>().First(l => l.Name == TableName);
        var inv = imported.Invoice;
        var t = inv.StatedTotals ?? InvoiceCalculator.Calculate(inv);

        var values = new object?[]
        {
            DateTime.Now,
            Path.GetFileName(path),
            imported.FormatDescription,
            inv.Number,
            inv.IssueDate,
            CodeLists.Display(CodeLists.InvoiceTypes, inv.TypeCode),
            inv.Seller.Name,
            inv.Seller.VatId,
            inv.BuyerReference,
            t.TaxBasisTotal,
            t.TaxTotal,
            t.GrandTotal,
            t.DuePayable,
            inv.CurrencyCode,
            inv.DueDate,
            inv.Payment.Iban,
            inv.Payment.RemittanceInformation,
            report.IsValid ? "gültig" : "ungültig",
            report.ErrorCount,
            report.WarningCount,
            path,
        };
        AppendRow(lo, values, report.IsValid);
    }

    /// <summary>Logs a file that could not be read at all.</summary>
    public static void AppendFailure(Excel.Workbook wb, string path, string error)
    {
        var ws = EnsureSheet(wb);
        var lo = ws.ListObjects.Cast<Excel.ListObject>().First(l => l.Name == TableName);
        var values = new object?[Columns.Length];
        values[0] = DateTime.Now;
        values[1] = Path.GetFileName(path);
        values[2] = "nicht lesbar: " + error;
        values[17] = "ungültig";
        values[20] = path;
        AppendRow(lo, values, false);
    }

    private static void AppendRow(Excel.ListObject lo, object?[] values, bool ok)
    {
        // A fresh table has one empty data row: reuse it instead of adding a second one.
        Excel.Range rowRange;
        var body = lo.DataBodyRange;
        if (lo.ListRows.Count == 1 && body is not null && App.WorksheetFunction.CountA(body) == 0)
            rowRange = body;
        else
            rowRange = lo.ListRows.Add().Range;

        var data = new object?[1, values.Length];
        for (var i = 0; i < values.Length; i++) data[0, i] = ToExcel(values[i]);
        rowRange.Value2 = data;
        ((Excel.Range)rowRange.Cells[1, 18]).Interior.Color = ok ? OkFill : ErrorFill;
    }
}
