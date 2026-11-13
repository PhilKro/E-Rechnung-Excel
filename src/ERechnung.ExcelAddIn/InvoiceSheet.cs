using ERechnung.Core.Calculation;
using ERechnung.Core.Codes;
using ERechnung.Core.Model;
using ERechnung.Core.Sheet;
using ERechnung.Core.Validation;
using Excel = Microsoft.Office.Interop.Excel;
using static ERechnung.ExcelAddIn.ExcelHelpers;

namespace ERechnung.ExcelAddIn;

/// <summary>
/// Renders <see cref="SheetLayout"/> as a worksheet and reads/writes invoice data through sheet-scoped names.
/// Layout: A = label, B = value (named ER_*), C = hint; totals in E/F; line table (ListObject ER_Positionen*) below.
/// </summary>
internal static class InvoiceSheet
{
    private const int ColLabel = 1, ColValue = 2, ColHint = 3, ColTotalsLabel = 5, ColTotalsValue = 6;
    private const int FirstRow = 4;
    private const int TotalsMaxRows = 22;

    /// <summary>Column widths A..J, shared by the header form and the line table.</summary>
    private static readonly double[] ColumnWidths = { 27, 38, 34, 13, 24, 16, 15, 26, 11, 15 };

    private static readonly string LinesHeaderTitle = "Positionen";

    public static bool IsInvoiceSheet(Excel.Worksheet? ws) => ws is not null && NamedRange(ws, SheetLayout.MarkerName) is not null;

    public static Excel.Worksheet? Active() => App.ActiveSheet as Excel.Worksheet;

    // ------------------------------------------------------------------ create

    /// <summary>Creates an empty invoice sheet with <paramref name="emptyLineRows"/> empty line rows.</summary>
    public static Excel.Worksheet Create(Excel.Workbook wb, string baseName, int emptyLineRows = 5)
    {
        ListsSheet.Ensure(wb);
        var ws = AddSheetAtEnd(wb, baseName);
        Batch(() => Render(wb, ws, Math.Max(1, emptyLineRows)));
        ws.Activate();
        Cell(ws, FirstRow + 1, ColValue).Select();
        return ws;
    }

    private static void Render(Excel.Workbook wb, Excel.Worksheet ws, int lineRows)
    {
        ws.Cells.Font.Name = "Calibri";
        ws.Cells.Font.Size = 10;
        for (var i = 0; i < ColumnWidths.Length; i++)
            ((Excel.Range)ws.Columns[i + 1]).ColumnWidth = ColumnWidths[i];

        var title = Cell(ws, 1, 1);
        title.Value2 = "E-Rechnung";
        title.Font.Size = 16;
        title.Font.Bold = true;
        var info = Cell(ws, 2, 1);
        info.Value2 = "Fett = Pflichtangabe für XRechnung. Gelbe Felder ausfüllen, Positionen unten eintragen. " +
                      "Summen werden beim Prüfen und Exportieren berechnet.";
        info.Font.Italic = true;
        info.Font.Color = HintColor;
        ws.Names.Add(Name: SheetLayout.MarkerName, RefersTo: title);

        var row = FirstRow;
        foreach (var section in SheetLayout.Sections)
        {
            SectionHeader(ws, row, section.Title, ColHint);
            row++;
            foreach (var f in section.Fields)
            {
                var label = Cell(ws, row, ColLabel);
                label.Value2 = f.Label;
                label.Font.Bold = f.Required;

                var value = Cell(ws, row, ColValue);
                value.Interior.Color = InputFill;
                value.Borders.Color = Rgb(210, 210, 210);
                value.NumberFormat = f.Kind switch
                {
                    FieldKind.Date => "dd.mm.yyyy",
                    FieldKind.Decimal => "#,##0.00",
                    _ => "@",                      // text: keeps leading zeros (PLZ), IBANs, "+49 …"
                };
                value.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (f.Codes is not null) AddListValidation(value, f.Codes);
                ws.Names.Add(Name: SheetLayout.NameFor(f.Key), RefersTo: value);

                var hint = Cell(ws, row, ColHint);
                hint.Value2 = f.Hint is null ? f.BusinessTerm : $"{f.BusinessTerm} · {f.Hint}";
                hint.Font.Color = HintColor;
                hint.Font.Size = 9;
                row++;
            }
            row++;
        }

        // Line table
        SectionHeader(ws, row, LinesHeaderTitle, SheetLayout.LineColumns.Count);
        row++;
        var cols = SheetLayout.LineColumns;
        var headerRange = Range(ws, row, 1, row, cols.Count);
        var headers = new object[1, cols.Count];
        for (var i = 0; i < cols.Count; i++) headers[0, i] = cols[i].Header;
        headerRange.Value2 = headers;

        var tableRange = Range(ws, row, 1, row + lineRows, cols.Count);
        var lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes);
        lo.Name = UniqueTableName(wb);
        lo.TableStyle = "TableStyleMedium2";
        lo.ShowTotals = false;
        for (var i = 0; i < cols.Count; i++)
        {
            var c = cols[i];
            var body = lo.ListColumns[i + 1].DataBodyRange;
            body.NumberFormat = c.Kind switch
            {
                FieldKind.Decimal when c.Key == LineKeys.Quantity => "#,##0.###",
                FieldKind.Decimal when c.Key == LineKeys.VatRate => "0.##",
                FieldKind.Decimal => "#,##0.00",
                _ => "@",
            };
            if (c.Codes is not null) AddListValidation(body, c.Codes);
        }
        SetNetAmountFormula(lo);

        // Defaults for empty rows: running position numbers, unit "Stück", 19 %
        var defaults = new object?[lineRows, cols.Count];
        for (var r = 0; r < lineRows; r++)
        {
            defaults[r, Index(LineKeys.Id)] = (r + 1).ToString();
            defaults[r, Index(LineKeys.Unit)] = CodeLists.Display(CodeLists.Units, "C62");
            defaults[r, Index(LineKeys.VatCategory)] = CodeLists.Display(CodeLists.VatCategories, "S");
            defaults[r, Index(LineKeys.VatRate)] = 19d;
        }
        WriteLineValues(lo, defaults, skipComputed: true);

        // Sensible defaults for header fields
        Set(ws, FieldKeys.TypeCode, CodeLists.Display(CodeLists.InvoiceTypes, "380"));
        Set(ws, FieldKeys.Currency, CodeLists.Display(CodeLists.Currencies, "EUR"));
        Set(ws, FieldKeys.SellerCountry, CodeLists.Display(CodeLists.Countries, "DE"));
        Set(ws, FieldKeys.BuyerCountry, CodeLists.Display(CodeLists.Countries, "DE"));
        Set(ws, FieldKeys.PaymentMeans, CodeLists.Display(CodeLists.PaymentMeans, "58"));
        Set(ws, FieldKeys.IssueDate, DateTime.Today);

        WriteTotalsPlaceholder(ws);

        var ps = ws.PageSetup;
        try
        {
            ps.Zoom = false;
            ps.FitToPagesWide = 1;
            ps.FitToPagesTall = false;
            ps.Orientation = Excel.XlPageOrientation.xlPortrait;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // No printer installed: page setup is optional.
        }
    }

    private static void SectionHeader(Excel.Worksheet ws, int row, string title, int width)
    {
        var r = Range(ws, row, 1, row, width);
        r.Interior.Color = SectionFill;
        r.Font.Color = Rgb(255, 255, 255);
        r.Font.Bold = true;
        Cell(ws, row, 1).Value2 = title;
    }

    private static void AddListValidation(Excel.Range target, IReadOnlyList<CodeEntry> codes)
    {
        target.Validation.Delete();
        target.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertWarning,
            Excel.XlFormatConditionOperator.xlBetween, "=" + ListsSheet.NameOf(codes));
        target.Validation.IgnoreBlank = true;
        target.Validation.InCellDropdown = true;
        target.Validation.ErrorMessage = "Dieser Code ist nicht in der Liste. Er wird trotzdem übernommen, wenn er gültig ist.";
    }

    private static void SetNetAmountFormula(Excel.ListObject lo)
    {
        var qty = Header(LineKeys.Quantity);
        var price = Header(LineKeys.UnitPrice);
        lo.ListColumns[Index(LineKeys.NetAmount) + 1].DataBodyRange.Formula =
            $"=ROUND([@[{qty}]]*[@[{price}]],2)";
    }

    private static string UniqueTableName(Excel.Workbook wb)
    {
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Excel.Worksheet w in wb.Worksheets)
            foreach (Excel.ListObject l in w.ListObjects) existing.Add(l.Name);
        for (var i = 1; ; i++)
        {
            var name = $"{SheetLayout.LineTablePrefix}{i}";
            if (!existing.Contains(name)) return name;
        }
    }

    private static int Index(string lineKey) => SheetLayout.LineColumns.ToList().FindIndex(c => c.Key == lineKey);

    private static string Header(string lineKey) => SheetLayout.LineColumns[Index(lineKey)].Header;

    public static Excel.ListObject? LineTable(Excel.Worksheet ws) =>
        ws.ListObjects.Cast<Excel.ListObject>().FirstOrDefault(l => l.Name.StartsWith(SheetLayout.LineTablePrefix, StringComparison.OrdinalIgnoreCase));

    // ------------------------------------------------------------------ read

    public static SheetReadResult Read(Excel.Worksheet ws)
    {
        var fields = new Dictionary<string, object?>();
        foreach (var f in SheetLayout.AllFields)
            fields[f.Key] = NamedRange(ws, SheetLayout.NameFor(f.Key))?.Value2;   // missing name (older sheet) = empty

        var rows = new List<IReadOnlyDictionary<string, object?>>();
        var lo = LineTable(ws);
        if (lo?.DataBodyRange is { } body)
        {
            // Map columns by header text, so users may reorder or add own columns.
            var headerCells = (object[,])lo.HeaderRowRange.Value2;
            var colIndex = new Dictionary<string, int>();
            for (var c = 1; c <= headerCells.GetLength(1); c++)
            {
                var h = headerCells[1, c]?.ToString()?.Trim();
                var def = SheetLayout.LineColumns.FirstOrDefault(d => string.Equals(d.Header, h, StringComparison.OrdinalIgnoreCase));
                if (def is not null) colIndex[def.Key] = c;
            }

            var values = body.Value2 is object[,] arr ? arr : new object[,] { { body.Value2 } };
            for (var r = values.GetLowerBound(0); r <= values.GetUpperBound(0); r++)
            {
                var row = new Dictionary<string, object?>();
                foreach (var (key, c) in colIndex)
                {
                    if (SheetLayout.LineColumns.First(d => d.Key == key).Computed) continue;
                    var v = values[r, c - 1 + values.GetLowerBound(1)];
                    row[key] = v is int err && err < 0 ? null : v;   // error values (#NV, …) arrive as negative ints
                }
                rows.Add(row);
            }
        }
        return SheetMapper.Read(fields, rows);
    }

    // ------------------------------------------------------------------ write

    /// <summary>Fills a freshly created sheet with an invoice (demo data or a received invoice).</summary>
    public static void Write(Excel.Worksheet ws, Invoice invoice, bool received, string? sourceInfo = null)
    {
        Batch(() =>
        {
            var (fields, lines) = SheetMapper.Write(invoice);
            foreach (var (key, value) in fields) Set(ws, key, value);

            var lo = LineTable(ws) ?? throw new InvalidOperationException("Positionstabelle nicht gefunden.");
            var cols = SheetLayout.LineColumns;
            var count = Math.Max(1, lines.Count);
            var header = lo.HeaderRowRange;
            lo.Resize(ws.Range[header.Cells[1, 1], header.Cells[1, cols.Count].Offset[count, 0]]);

            var data = new object?[count, cols.Count];
            for (var r = 0; r < lines.Count; r++)
                for (var c = 0; c < cols.Count; c++)
                    data[r, c] = ToExcel(lines[r].GetValueOrDefault(cols[c].Key));
            if (lines.Count == 0) data[0, Index(LineKeys.Id)] = "1";

            // Received invoices show the line amounts as stated in the document (no formula),
            // so differences to quantity × price stay visible instead of being "corrected".
            WriteLineValues(lo, data, skipComputed: !received);
            if (!received) SetNetAmountFormula(lo);

            if (received)
            {
                var info = Cell(ws, 2, 1);
                info.Value2 = "Empfangene E-Rechnung" + (sourceInfo is null ? "" : $": {sourceInfo}") +
                              " – Summen laut Rechnung. Prüfergebnis im Blatt \"Prüfbericht\".";
                Cell(ws, 1, 1).Value2 = "E-Rechnung (Eingang)";
            }
            WriteTotals(ws, received && invoice.StatedTotals is not null ? invoice.StatedTotals : InvoiceCalculator.Calculate(invoice), received);
        });
    }

    private static void WriteLineValues(Excel.ListObject lo, object?[,] data, bool skipComputed)
    {
        var cols = SheetLayout.LineColumns;
        var rows = data.GetLength(0);
        for (var c = 0; c < cols.Count; c++)
        {
            if (skipComputed && cols[c].Computed) continue;
            var column = new object?[rows, 1];
            for (var r = 0; r < rows; r++) column[r, 0] = data[r, c];
            var body = lo.ListColumns[c + 1].DataBodyRange;
            ((Excel.Range)body.Resize[rows, 1]).Value2 = column;
        }
    }

    public static void Set(Excel.Worksheet ws, string key, object? value)
    {
        var cell = NamedRange(ws, SheetLayout.NameFor(key));
        if (cell is not null) cell.Value2 = ToExcel(value);
    }

    // ------------------------------------------------------------------ totals

    private static void WriteTotalsPlaceholder(Excel.Worksheet ws)
    {
        ClearTotals(ws);
        SectionHeader2(ws, FirstRow, "Summen");
        var c = Cell(ws, FirstRow + 1, ColTotalsLabel);
        c.Value2 = "werden beim Prüfen berechnet";
        c.Font.Italic = true;
        c.Font.Color = HintColor;
    }

    private static void SectionHeader2(Excel.Worksheet ws, int row, string title)
    {
        var r = Range(ws, row, ColTotalsLabel, row, ColTotalsValue);
        r.Interior.Color = SectionFill;
        r.Font.Color = Rgb(255, 255, 255);
        r.Font.Bold = true;
        Cell(ws, row, ColTotalsLabel).Value2 = title;
    }

    private static void ClearTotals(Excel.Worksheet ws)
    {
        var r = Range(ws, FirstRow, ColTotalsLabel, FirstRow + TotalsMaxRows, ColTotalsValue);
        r.Clear();
    }

    /// <summary>Writes totals and VAT breakdown into E/F (values, not formulas: they come from the calculator or the document).</summary>
    public static void WriteTotals(Excel.Worksheet ws, InvoiceTotals t, bool stated)
    {
        ClearTotals(ws);
        SectionHeader2(ws, FirstRow, stated ? "Summen (laut Rechnung)" : "Summen (berechnet)");
        var row = FirstRow + 1;

        void Line(string label, decimal amount, bool bold = false)
        {
            if (row > FirstRow + TotalsMaxRows) return;
            var l = Cell(ws, row, ColTotalsLabel);
            l.Value2 = label;
            var v = Cell(ws, row, ColTotalsValue);
            v.Value2 = (double)amount;
            v.NumberFormat = "#,##0.00";
            l.Font.Bold = bold;
            v.Font.Bold = bold;
            row++;
        }

        Line("Summe Positionen netto", t.LineTotal);
        if (t.AllowanceTotal != 0) Line("Abschläge", -t.AllowanceTotal);
        if (t.ChargeTotal != 0) Line("Zuschläge", t.ChargeTotal);
        Line("Rechnungsbetrag netto", t.TaxBasisTotal);
        foreach (var v in t.VatBreakdown)
        {
            var cat = CodeLists.LabelOf(CodeLists.VatCategories, v.Category) is { } label && v.Category != "S"
                ? $" ({v.Category})" : "";
            Line($"USt {v.Rate:0.##} %{cat} auf {v.TaxableAmount:N2}", v.TaxAmount);
        }
        Line("Umsatzsteuer gesamt", t.TaxTotal);
        Line("Rechnungsbetrag brutto", t.GrandTotal, bold: true);
        if (t.PrepaidAmount != 0) Line("Bereits gezahlt", -t.PrepaidAmount);
        if (t.RoundingAmount != 0) Line("Rundung", t.RoundingAmount);
        Line("Zahlbetrag", t.DuePayable, bold: true);
    }

    // ------------------------------------------------------------------ highlighting

    /// <summary>Colours value cells with errors (red) or warnings (yellow); resets all other input cells.</summary>
    public static void Highlight(Excel.Worksheet ws, ValidationReport report)
    {
        Batch(() =>
        {
            foreach (var f in SheetLayout.AllFields)
                if (NamedRange(ws, SheetLayout.NameFor(f.Key)) is { } cell) cell.Interior.Color = InputFill;

            var lo = LineTable(ws);
            if (lo is not null) lo.HeaderRowRange.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

            foreach (var m in report.Messages.Where(m => m.FieldKey is not null).OrderBy(m => m.Severity == Severity.Error ? 1 : 0))
            {
                var color = m.Severity == Severity.Error ? ErrorFill : m.Severity == Severity.Warning ? WarningFill : (int?)null;
                if (color is null) continue;
                if (m.FieldKey == FieldKeys.Lines)
                {
                    if (lo is not null) lo.HeaderRowRange.Interior.Color = color.Value;
                }
                else if (NamedRange(ws, SheetLayout.NameFor(m.FieldKey!)) is { } cell)
                {
                    cell.Interior.Color = color.Value;
                }
            }
        });
    }

    /// <summary>"'Rechnung'!B5" for hyperlinks from the report to a field.</summary>
    public static string? AddressOf(Excel.Worksheet ws, string fieldKey)
    {
        if (fieldKey == FieldKeys.Lines)
            return LineTable(ws) is { } lo ? $"'{ws.Name}'!{lo.HeaderRowRange.Address[false, false]}" : null;
        var cell = NamedRange(ws, SheetLayout.NameFor(fieldKey));
        return cell is null ? null : $"'{ws.Name}'!{cell.Address[false, false]}";
    }

    // ------------------------------------------------------------------ PDF

    /// <summary>
    /// Exports the sheet as PDF (the visual part of a ZUGFeRD invoice). The hint column and the input colours are
    /// hidden during export so the PDF looks like an invoice rather than a form.
    /// </summary>
    public static void ExportVisualPdf(Excel.Worksheet ws, string pdfPath)
    {
        var hintColumn = (Excel.Range)ws.Columns[ColHint];
        var infoRow = (Excel.Range)ws.Rows[2];
        var hintHidden = hintColumn.Hidden;
        var infoHidden = infoRow.Hidden;
        var app = App;
        app.ScreenUpdating = false;
        try
        {
            hintColumn.Hidden = true;
            infoRow.Hidden = true;
            ws.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath, Excel.XlFixedFormatQuality.xlQualityStandard,
                IncludeDocProperties: true, IgnorePrintAreas: false, OpenAfterPublish: false);
        }
        finally
        {
            hintColumn.Hidden = hintHidden;
            infoRow.Hidden = infoHidden;
            app.ScreenUpdating = true;
        }
    }
}
