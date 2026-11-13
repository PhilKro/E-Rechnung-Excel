using ERechnung.Core.Codes;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ERechnung.ExcelAddIn;

internal static class ExcelHelpers
{
    public static Excel.Application App => (Excel.Application)ExcelDnaUtil.Application;

    /// <summary>Excel colours are BGR integers.</summary>
    public static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

    public static readonly int InputFill = Rgb(255, 252, 230);
    public static readonly int ErrorFill = Rgb(255, 199, 206);
    public static readonly int WarningFill = Rgb(255, 235, 156);
    public static readonly int SectionFill = Rgb(31, 78, 121);
    public static readonly int HintColor = Rgb(118, 118, 118);
    public static readonly int OkFill = Rgb(198, 239, 206);

    public static Excel.Range Cell(Excel.Worksheet ws, int row, int col) => (Excel.Range)ws.Cells[row, col];

    public static Excel.Range Range(Excel.Worksheet ws, int r1, int c1, int r2, int c2) =>
        ws.Range[Cell(ws, r1, c1), Cell(ws, r2, c2)];

    /// <summary>Returns a sheet name that does not exist yet: "Rechnung", "Rechnung (2)", …</summary>
    public static string UniqueSheetName(Excel.Workbook wb, string baseName)
    {
        var clean = new string(baseName.Where(c => !"[]:*?/\\".Contains(c)).ToArray()).Trim();
        if (clean.Length == 0) clean = "Rechnung";
        if (clean.Length > 31) clean = clean[..31];
        var existing = wb.Worksheets.Cast<Excel.Worksheet>().Select(w => w.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!existing.Contains(clean)) return clean;
        for (var i = 2; ; i++)
        {
            var suffix = $" ({i})";
            var candidate = (clean.Length + suffix.Length > 31 ? clean[..(31 - suffix.Length)] : clean) + suffix;
            if (!existing.Contains(candidate)) return candidate;
        }
    }

    public static Excel.Worksheet? FindSheet(Excel.Workbook wb, string name) =>
        wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Sheet-scoped name lookup that returns null instead of throwing.</summary>
    public static Excel.Range? NamedRange(Excel.Worksheet ws, string name)
    {
        foreach (Excel.Name n in ws.Names)
        {
            // Sheet-scoped names are reported as "Sheet!Name"
            var local = n.Name.Contains('!') ? n.Name[(n.Name.IndexOf('!') + 1)..] : n.Name;
            if (string.Equals(local, name, StringComparison.OrdinalIgnoreCase))
            {
                try { return n.RefersToRange; }
                catch (System.Runtime.InteropServices.COMException) { return null; }   // #REF! after deleting cells
            }
        }
        return null;
    }

    public static Excel.Worksheet AddSheetAtEnd(Excel.Workbook wb, string baseName)
    {
        var last = wb.Worksheets[wb.Worksheets.Count];
        var ws = (Excel.Worksheet)wb.Worksheets.Add(After: last);
        ws.Name = UniqueSheetName(wb, baseName);
        return ws;
    }

    /// <summary>Converts a .NET value to what Excel's Value2 accepts (dates as OLE automation doubles).</summary>
    public static object? ToExcel(object? v) => v switch
    {
        null => null,
        DateTime dt => dt.ToOADate(),
        decimal m => (double)m,
        string s when s.Length == 0 => null,
        _ => v,
    };

    /// <summary>Runs an action with wait cursor and status bar text, restoring both afterwards.</summary>
    public static T Busy<T>(string status, Func<T> action)
    {
        var app = App;
        var oldCursor = app.Cursor;
        app.Cursor = Excel.XlMousePointer.xlWait;
        app.StatusBar = status;
        try
        {
            return action();
        }
        finally
        {
            app.Cursor = oldCursor;
            app.StatusBar = false;
        }
    }

    public static void Busy(string status, Action action) => Busy(status, () => { action(); return 0; });

    /// <summary>Faster bulk edits: no repaint, no recalculation, no events while building sheets.</summary>
    public static void Batch(Action action)
    {
        var app = App;
        var screen = app.ScreenUpdating;
        var events = app.EnableEvents;
        var calc = app.Calculation;
        app.ScreenUpdating = false;
        app.EnableEvents = false;
        try
        {
            try { app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch (System.Runtime.InteropServices.COMException) { }
            action();
        }
        finally
        {
            try { app.Calculation = calc; } catch (System.Runtime.InteropServices.COMException) { }
            app.EnableEvents = events;
            app.ScreenUpdating = screen;
        }
    }
}

/// <summary>
/// Hidden sheet "ER_Listen" holding the code lists used by drop-downs. List validation with literal values
/// would break on commas and the 255-character limit, so each list lives in a range with a workbook name.
/// </summary>
internal static class ListsSheet
{
    public const string SheetName = "ER_Listen";

    public static readonly (string Name, IReadOnlyList<CodeEntry> Entries)[] Lists =
    {
        ("InvoiceTypes", CodeLists.InvoiceTypes),
        ("Currencies", CodeLists.Currencies),
        ("Countries", CodeLists.Countries),
        ("PaymentMeans", CodeLists.PaymentMeans),
        ("Units", CodeLists.Units),
        ("VatCategories", CodeLists.VatCategories),
        ("ElectronicAddressSchemes", CodeLists.ElectronicAddressSchemes),
    };

    public static string NameOf(IReadOnlyList<CodeEntry> list)
    {
        var hit = Lists.FirstOrDefault(l => ReferenceEquals(l.Entries, list));
        return hit.Name is null ? throw new ArgumentException("Unknown code list") : "ER_Liste_" + hit.Name;
    }

    /// <summary>Creates or refreshes the list sheet (refreshing keeps lists current after add-in updates).</summary>
    public static void Ensure(Excel.Workbook wb)
    {
        var active = wb.Application.ActiveSheet as Excel.Worksheet;
        var ws = ExcelHelpers.FindSheet(wb, SheetName);
        if (ws is null)
        {
            ws = ExcelHelpers.AddSheetAtEnd(wb, SheetName);
            ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }

        ws.Cells.ClearContents();
        for (var c = 0; c < Lists.Length; c++)
        {
            var (name, entries) = Lists[c];
            ExcelHelpers.Cell(ws, 1, c + 1).Value2 = name;
            var values = new object[entries.Count, 1];
            for (var i = 0; i < entries.Count; i++) values[i, 0] = entries[i].Display;
            var range = ExcelHelpers.Range(ws, 2, c + 1, entries.Count + 1, c + 1);
            range.NumberFormat = "@";
            range.Value2 = values;

            var fullName = "ER_Liste_" + name;
            try { wb.Names.Item(fullName).Delete(); } catch (System.Runtime.InteropServices.COMException) { } catch (ArgumentException) { }
            wb.Names.Add(Name: fullName, RefersTo: range);
        }
        active?.Activate();
    }
}
