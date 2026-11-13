using System.Diagnostics;
using System.Reflection;
using ERechnung.Core;
using ERechnung.Core.Calculation;
using ERechnung.Core.IO;
using ERechnung.Core.Model;
using ERechnung.Core.Samples;
using ERechnung.Core.Sheet;
using ERechnung.Core.Validation;
using Excel = Microsoft.Office.Interop.Excel;
using static ERechnung.ExcelAddIn.ExcelHelpers;

namespace ERechnung.ExcelAddIn;

public enum ExportKind
{
    XRechnungUbl,
    XRechnungCii,
    ZugferdPdf,
}

/// <summary>The actions behind the ribbon buttons. Every public method is safe to call from a ribbon callback.</summary>
internal static class Commands
{
    /// <summary>Sheet-scoped name marking a sheet created from a received invoice.</summary>
    private const string ReceivedMarker = SheetLayout.NamePrefix + "Empfangen";

    // ------------------------------------------------------------------ create

    public static void NewInvoice()
    {
        var wb = App.ActiveWorkbook ?? App.Workbooks.Add();
        InvoiceSheet.Create(wb, "Rechnung");
    }

    public static void InsertDemo()
    {
        var wb = App.ActiveWorkbook ?? App.Workbooks.Add();
        var ws = InvoiceSheet.Create(wb, "Beispielrechnung");
        InvoiceSheet.Write(ws, DemoInvoices.Standard(), received: false);
        ws.Activate();
    }

    // ------------------------------------------------------------------ check

    public static void Validate()
    {
        var ws = RequireInvoiceSheet();
        if (ws is null) return;
        var wb = (Excel.Workbook)ws.Parent;

        var (invoice, report) = CheckSheet(ws, ValidationProfile.XRechnung);
        if (!IsReceived(ws))
        {
            // Check what an export would produce: the official validator only sees XML.
            try
            {
                RunKosit(report, InvoiceExporter.ToXRechnungXml(invoice, XmlSyntax.Ubl), $"{invoice.Number ?? "rechnung"}.xml");
            }
            catch (InvoiceDataException ex)
            {
                report.Add(Severity.Error, "EINGABE", ex.Message);
            }
        }
        InvoiceSheet.Highlight(ws, report);
        ReportSheet.Show(wb, report, $"Blatt \"{ws.Name}\" (geprüft als XRechnung 3.0)", ws);
        ws.Activate();
        App.StatusBar = false;
        (report.IsValid ? (Action<string>)Ui.Info : Ui.Warn)(report.Summary + "\n\nDetails im Blatt \"Prüfbericht\".");
    }

    // ------------------------------------------------------------------ export

    public static void Export(ExportKind kind)
    {
        var ws = RequireInvoiceSheet();
        if (ws is null) return;
        var wb = (Excel.Workbook)ws.Parent;

        var profile = kind == ExportKind.ZugferdPdf ? ValidationProfile.EN16931 : ValidationProfile.XRechnung;
        var (invoice, report) = CheckSheet(ws, profile);
        InvoiceSheet.Highlight(ws, report);
        var title = $"{invoice.Number} ({Describe(kind)})";

        if (report.ErrorCount > 0)
        {
            ReportSheet.Show(wb, report, title, ws);
            if (!Ui.Confirm($"Die Rechnung hat {report.ErrorCount} Fehler (siehe Blatt \"Prüfbericht\").\n" +
                            "Eine fehlerhafte E-Rechnung kann vom Empfänger zurückgewiesen werden.\n\nTrotzdem exportieren?"))
                return;
        }

        byte[] fileBytes, xml;
        try
        {
            switch (kind)
            {
                case ExportKind.XRechnungUbl:
                    fileBytes = xml = InvoiceExporter.ToXRechnungXml(invoice, XmlSyntax.Ubl);
                    break;
                case ExportKind.XRechnungCii:
                    fileBytes = xml = InvoiceExporter.ToXRechnungXml(invoice, XmlSyntax.Cii);
                    break;
                default:
                    fileBytes = BuildZugferdPdf(ws, invoice);
                    xml = InvoiceImporter.Import(fileBytes).Xml;   // validate exactly what is embedded
                    break;
            }
        }
        catch (InvoiceDataException ex)
        {
            Ui.Error("Die Rechnung kann nicht erzeugt werden:\n\n" + ex.Message);
            return;
        }

        RunKosit(report, xml, title);
        if (report.KositAccepted == false)
        {
            ReportSheet.Show(wb, report, title, ws);
            if (!Ui.Confirm("Der offizielle KoSIT-Validator hat die Rechnung abgelehnt (siehe Blatt \"Prüfbericht\").\n\nTrotzdem speichern?"))
                return;
        }

        var (ext, filter) = kind == ExportKind.ZugferdPdf
            ? (".pdf", "ZUGFeRD-PDF (*.pdf)|*.pdf")
            : (".xml", "XRechnung (*.xml)|*.xml");
        var suffix = kind switch
        {
            ExportKind.XRechnungUbl => "_XRechnung",
            ExportKind.XRechnungCii => "_XRechnung_CII",
            _ => "",
        };
        var defaultName = KositValidator.SafeFileName((invoice.Number ?? "Rechnung") + suffix) + ext;
        var path = Ui.SaveFile($"{Describe(kind)} speichern", filter, defaultName, InitialDir(wb));
        if (path is null) return;

        File.WriteAllBytes(path, fileBytes);
        ReportSheet.Show(wb, report, $"{Path.GetFileName(path)} ({Describe(kind)})", ws);
        ws.Activate();
        Ui.Info($"Gespeichert:\n{path}\n\n{report.Summary}");
    }

    private static byte[] BuildZugferdPdf(Excel.Worksheet ws, Invoice invoice)
    {
        var tmp = Path.Combine(Path.GetTempPath(), "ERechnungExcel");
        Directory.CreateDirectory(tmp);
        var visualPdf = Path.Combine(tmp, $"visual-{Guid.NewGuid():N}.pdf");
        try
        {
            Busy("PDF wird erzeugt …", () => InvoiceSheet.ExportVisualPdf(ws, visualPdf));
            using var visual = File.OpenRead(visualPdf);
            using var output = new MemoryStream();
            InvoiceExporter.WriteZugferdPdf(invoice, visual, output, ZugferdProfile.EN16931);
            return output.ToArray();
        }
        finally
        {
            try { File.Delete(visualPdf); } catch (IOException) { }
        }
    }

    // ------------------------------------------------------------------ receive

    public static void Import()
    {
        var files = Ui.OpenFiles("E-Rechnungen öffnen", "E-Rechnungen (*.xml;*.pdf)|*.xml;*.pdf|Alle Dateien (*.*)|*.*",
            App.ActiveWorkbook is { } active ? InitialDir(active) : null);
        if (files.Length == 0) return;
        var wb = App.ActiveWorkbook ?? App.Workbooks.Add();

        var imported = new List<(string Path, ImportedInvoice Doc)>();
        var failures = new List<(string Path, string Error)>();
        Busy("E-Rechnungen werden gelesen …", () =>
        {
            foreach (var f in files)
            {
                try { imported.Add((f, InvoiceImporter.Import(f))); }
                catch (InvoiceDataException ex) { failures.Add((f, ex.Message)); }
            }
        });

        // Built-in checks with the profile the document claims (XRechnung or plain EN 16931)
        var reports = imported.Select(i =>
        {
            var profile = i.Doc.SpecificationId?.Contains("xrechnung", StringComparison.OrdinalIgnoreCase) == true
                ? ValidationProfile.XRechnung : ValidationProfile.EN16931;
            return BuiltInValidator.Validate(i.Doc.Invoice, profile);
        }).ToList();

        // One KoSIT run for all files (JVM start dominates)
        var kosit = Settings.Current.TryGetKosit(out var reason);
        if (kosit is null || imported.Count == 0)
        {
            foreach (var r in reports) r.KositSkippedReason = reason;
        }
        else
        {
            try
            {
                var results = Busy($"KoSIT prüft {imported.Count} Rechnung(en) …",
                    () => kosit.ValidateMany(imported.Select(i => (Path.GetFileName(i.Path), i.Doc.Xml)).ToList()));
                for (var i = 0; i < results.Count; i++) Merge(reports[i], results[i]);
            }
            catch (Exception ex) when (ex is InvalidOperationException or TimeoutException or IOException)
            {
                foreach (var r in reports) r.KositSkippedReason = "Fehler: " + ex.Message;
            }
        }

        Batch(() =>
        {
            for (var i = 0; i < imported.Count; i++) InboxSheet.Append(wb, imported[i].Path, imported[i].Doc, reports[i]);
            foreach (var (path, error) in failures) InboxSheet.AppendFailure(wb, path, error);
        });

        if (imported.Count == 1)
        {
            var (path, doc) = imported[0];
            var ws = InvoiceSheet.Create(wb, "Eingang " + (doc.Invoice.Number ?? Path.GetFileNameWithoutExtension(path)), emptyLineRows: 1);
            ws.Names.Add(Name: ReceivedMarker, RefersTo: Cell(ws, 1, 1));
            InvoiceSheet.Write(ws, doc.Invoice, received: true, $"{Path.GetFileName(path)} ({doc.FormatDescription})");
            InvoiceSheet.Highlight(ws, reports[0]);
            ReportSheet.Show(wb, reports[0], $"{Path.GetFileName(path)} ({doc.FormatDescription})", ws);
            ws.Activate();
            (reports[0].IsValid ? (Action<string>)Ui.Info : Ui.Warn)(
                $"{Path.GetFileName(path)}\n{doc.FormatDescription}\n\n{reports[0].Summary}\n\nDetails im Blatt \"Prüfbericht\", Eintrag im Blatt \"{InboxSheet.SheetName}\".");
        }
        else
        {
            InboxSheet.EnsureSheet(wb).Activate();
            var invalid = reports.Count(r => !r.IsValid);
            var msg = $"{imported.Count} E-Rechnung(en) eingelesen, davon {invalid} ungültig.";
            if (failures.Count > 0)
                msg += $"\n\n{failures.Count} Datei(en) nicht lesbar:\n" +
                       string.Join("\n", failures.Take(10).Select(f => $"• {Path.GetFileName(f.Path)}: {f.Error}"));
            msg += $"\n\nÜbersicht im Blatt \"{InboxSheet.SheetName}\". Einzelne Rechnung im Detail: Datei einzeln öffnen.";
            (invalid == 0 && failures.Count == 0 ? (Action<string>)Ui.Info : Ui.Warn)(msg);
        }

        if (imported.Count == 0 && failures.Count == 1) Ui.Error($"{Path.GetFileName(failures[0].Path)}:\n\n{failures[0].Error}");
    }

    public static void ShowInbox()
    {
        var wb = App.ActiveWorkbook ?? App.Workbooks.Add();
        InboxSheet.EnsureSheet(wb).Activate();
    }

    // ------------------------------------------------------------------ KoSIT

    public static void KositStatus()
    {
        var settings = KositSettings.FromInstallDir(Settings.Current.KositDir);
        var problem = settings.FindProblem();
        if (problem is null)
        {
            var versions = Path.Combine(settings.InstallDir, "versions.txt");
            var info = File.Exists(versions) ? File.ReadAllText(versions).Trim() : "(keine Versionsinfo)";
            Ui.Info($"Der offizielle KoSIT-Validator ist installiert:\n{settings.InstallDir}\n\n{info}\n\n" +
                    $"Prüfung mit KoSIT ist {(Settings.Current.UseKosit ? "eingeschaltet" : "ausgeschaltet")}.\n\n" +
                    "Zum Aktualisieren erneut \"KoSIT installieren\" wählen.");
            return;
        }
        if (Ui.Confirm($"Der KoSIT-Validator ist nicht installiert ({problem}).\n\n" +
                       "Ohne ihn prüft das Add-in nur die wichtigsten Regeln selbst. Der KoSIT-Validator prüft vollständig nach den " +
                       "offiziellen Regeln (wie die Rechnungsportale von Bund und Ländern).\n\n" +
                       "Jetzt installieren? (Download ca. 60 MB von GitHub und adoptium.net, keine Administratorrechte nötig)"))
            InstallKosit();
    }

    public static void InstallKosit()
    {
        var tmp = Path.Combine(Path.GetTempPath(), "ERechnungExcel");
        Directory.CreateDirectory(tmp);
        var script = Path.Combine(tmp, "setup-kosit.ps1");
        using (var res = Assembly.GetExecutingAssembly().GetManifestResourceStream("setup-kosit.ps1")
                         ?? throw new InvalidOperationException("Installationsskript fehlt im Add-in."))
        using (var file = File.Create(script))
            res.CopyTo(file);

        var args = $"-NoProfile -ExecutionPolicy Bypass -NoExit -File \"{script}\"";
        if (!string.IsNullOrWhiteSpace(Settings.Current.KositDir)) args += $" -TargetDir \"{Settings.Current.KositDir}\"";
        Process.Start(new ProcessStartInfo("powershell.exe", args) { UseShellExecute = true });
        Ui.Info("Die Installation läuft im PowerShell-Fenster. Wenn dort \"KoSIT validator installed\" steht, " +
                "kann das Fenster geschlossen werden – die Prüfung ist dann sofort aktiv.");
    }

    public static void SetKositEnabled(bool enabled)
    {
        Settings.Current.UseKosit = enabled;
        Settings.Current.Save();
    }

    // ------------------------------------------------------------------ helpers

    private static Excel.Worksheet? RequireInvoiceSheet()
    {
        var ws = InvoiceSheet.Active();
        if (InvoiceSheet.IsInvoiceSheet(ws)) return ws;
        Ui.Warn("Das aktive Blatt ist kein Rechnungsblatt.\n\nBitte ein mit \"Neue Rechnung\" erstelltes Blatt auswählen " +
                "oder eine neue Rechnung anlegen.");
        return null;
    }

    private static bool IsReceived(Excel.Worksheet ws) => NamedRange(ws, ReceivedMarker) is not null;

    /// <summary>Reads the sheet, runs the built-in checks and refreshes the totals area.</summary>
    private static (Invoice Invoice, ValidationReport Report) CheckSheet(Excel.Worksheet ws, ValidationProfile profile)
    {
        var read = InvoiceSheet.Read(ws);
        var report = new ValidationReport();
        report.Messages.AddRange(read.Problems);
        BuiltInValidator.Validate(read.Invoice, profile, report);
        if (!IsReceived(ws))
            Batch(() => InvoiceSheet.WriteTotals(ws, InvoiceCalculator.Calculate(read.Invoice), stated: false));
        return (read.Invoice, report);
    }

    private static void RunKosit(ValidationReport report, byte[] xml, string documentName)
    {
        var kosit = Settings.Current.TryGetKosit(out var reason);
        if (kosit is null)
        {
            report.KositSkippedReason = reason;
            return;
        }
        try
        {
            var result = Busy("Prüfung mit dem KoSIT-Validator läuft …", () => kosit.Validate(xml, documentName));
            Merge(report, result);
        }
        catch (Exception ex) when (ex is InvalidOperationException or TimeoutException or IOException)
        {
            report.KositSkippedReason = "Fehler: " + ex.Message;
        }
    }

    /// <summary>Adds KoSIT findings; a rule the built-in check already reported (with a link to the field) is not repeated.</summary>
    private static void Merge(ValidationReport report, KositResult result)
    {
        report.KositAccepted = result.Accepted;
        report.KositScenario = result.Scenario;
        report.KositHtmlReport = result.HtmlReport;
        var known = report.Messages.Select(m => m.Code).ToHashSet(StringComparer.OrdinalIgnoreCase);
        report.Messages.AddRange(result.Messages.Where(m => !known.Contains(m.Code)));
    }

    private static string InitialDir(Excel.Workbook wb) =>
        !string.IsNullOrEmpty(wb.Path) ? wb.Path : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    private static string Describe(ExportKind kind) => kind switch
    {
        ExportKind.XRechnungUbl => "XRechnung UBL",
        ExportKind.XRechnungCii => "XRechnung CII",
        _ => "ZUGFeRD-PDF",
    };
}
