using ERechnung.Core.Sheet;
using ERechnung.Core.Validation;
using Excel = Microsoft.Office.Interop.Excel;
using static ERechnung.ExcelAddIn.ExcelHelpers;

namespace ERechnung.ExcelAddIn;

/// <summary>The "Prüfbericht" sheet: one row per validation message, with links to the affected field.</summary>
internal static class ReportSheet
{
    public const string SheetName = "Prüfbericht";

    public static void Show(Excel.Workbook wb, ValidationReport report, string documentTitle, Excel.Worksheet? invoiceSheet)
    {
        var ws = FindSheet(wb, SheetName) ?? AddSheetAtEnd(wb, SheetName);

        Batch(() =>
        {
            ws.Cells.Clear();
            foreach (Excel.Hyperlink h in ws.Hyperlinks) h.Delete();
            ws.Cells.Font.Name = "Calibri";
            ws.Cells.Font.Size = 10;
            double[] widths = { 10, 10, 22, 90, 28 };
            for (var i = 0; i < widths.Length; i++) ((Excel.Range)ws.Columns[i + 1]).ColumnWidth = widths[i];

            var title = Cell(ws, 1, 1);
            title.Value2 = "Prüfbericht";
            title.Font.Size = 16;
            title.Font.Bold = true;

            Cell(ws, 2, 1).Value2 = "Dokument:";
            Cell(ws, 2, 3).Value2 = documentTitle;
            Cell(ws, 3, 1).Value2 = "Geprüft am:";
            Cell(ws, 3, 3).Value2 = DateTime.Now.ToOADate();
            Cell(ws, 3, 3).NumberFormat = "dd.mm.yyyy hh:mm";
            Cell(ws, 3, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            Cell(ws, 4, 1).Value2 = "Ergebnis:";
            var verdict = Cell(ws, 4, 3);
            verdict.Value2 = report.Summary;
            verdict.Font.Bold = true;
            Range(ws, 4, 3, 4, 4).Interior.Color = report.IsValid ? OkFill : ErrorFill;

            var row = 5;
            if (report.KositScenario is not null)
            {
                Cell(ws, row, 1).Value2 = "KoSIT-Szenario:";
                Cell(ws, row, 3).Value2 = report.KositScenario;
                row++;
            }
            if (report.KositHtmlReport is not null && File.Exists(report.KositHtmlReport))
            {
                Cell(ws, row, 1).Value2 = "KoSIT-Bericht:";
                ws.Hyperlinks.Add(Cell(ws, row, 3), report.KositHtmlReport, Type.Missing, "Offiziellen Prüfbericht im Browser öffnen",
                    "Offiziellen KoSIT-Prüfbericht öffnen (HTML)");
                row++;
            }

            row++;
            string[] headers = { "Schwere", "Quelle", "Regel", "Meldung", "Feld / Ort" };
            for (var c = 0; c < headers.Length; c++) Cell(ws, row, c + 1).Value2 = headers[c];
            var hdr = Range(ws, row, 1, row, headers.Length);
            hdr.Font.Bold = true;
            hdr.Interior.Color = SectionFill;
            hdr.Font.Color = Rgb(255, 255, 255);
            row++;

            if (report.Messages.Count == 0)
            {
                Cell(ws, row, 4).Value2 = "Keine Meldungen.";
                row++;
            }

            foreach (var m in report.Messages.OrderBy(m => m.Severity).ThenBy(m => m.Source))
            {
                Cell(ws, row, 1).Value2 = m.Severity switch
                {
                    Severity.Error => "Fehler",
                    Severity.Warning => "Warnung",
                    _ => "Hinweis",
                };
                Range(ws, row, 1, row, 1).Interior.Color = m.Severity switch
                {
                    Severity.Error => ErrorFill,
                    Severity.Warning => WarningFill,
                    _ => OkFill,
                };
                Cell(ws, row, 2).Value2 = m.Source == MessageSource.Kosit ? "KoSIT" : "Add-in";
                Cell(ws, row, 3).Value2 = m.Code;
                var text = Cell(ws, row, 4);
                text.Value2 = m.Text;
                text.WrapText = true;

                var target = Cell(ws, row, 5);
                var link = invoiceSheet is not null && m.FieldKey is not null ? InvoiceSheet.AddressOf(invoiceSheet, m.FieldKey) : null;
                var label = m.FieldKey is null ? m.Location : FieldLabel(m.FieldKey);
                if (link is not null)
                    ws.Hyperlinks.Add(target, "", link, "Zum Feld springen", label ?? "Feld");
                else if (label is not null)
                    target.Value2 = ShortLocation(label);
                row++;
            }
        });

        ws.Activate();
    }

    private static string FieldLabel(string key)
    {
        if (key == FieldKeys.Lines) return "Positionen";
        var f = SheetLayout.Field(key);
        if (f is null) return key;
        var section = SheetLayout.Sections.First(s => s.Fields.Contains(f)).Title.Split(' ')[0];
        return $"{section}: {f.Label}";
    }

    /// <summary>KoSIT XPaths are long and namespace-qualified; keep only local element names.</summary>
    private static string ShortLocation(string xpath)
    {
        var parts = xpath.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Contains('}') ? p[(p.IndexOf('}') + 1)..] : p);
        var s = string.Join("/", parts);
        return s.Length > 250 ? "…" + s[^250..] : s;
    }
}
