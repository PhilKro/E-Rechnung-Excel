using ExcelDna.Integration;

namespace ERechnung.ExcelAddIn;

/// <summary>
/// The ribbon actions as Excel macros, callable via Application.Run("ERechnung_…") or assignable to buttons
/// and keyboard shortcuts. The smoke test (tests/excel-smoke.ps1) drives Excel through these.
/// </summary>
public static class Macros
{
    [ExcelCommand(Name = "ERechnung_Neu")] public static void New() => Run(Commands.NewInvoice);
    [ExcelCommand(Name = "ERechnung_Beispiel")] public static void Demo() => Run(Commands.InsertDemo);
    [ExcelCommand(Name = "ERechnung_Pruefen")] public static void Check() => Run(Commands.Validate);
    [ExcelCommand(Name = "ERechnung_ExportUBL")] public static void ExportUbl() => Run(() => Commands.Export(ExportKind.XRechnungUbl));
    [ExcelCommand(Name = "ERechnung_ExportCII")] public static void ExportCii() => Run(() => Commands.Export(ExportKind.XRechnungCii));
    [ExcelCommand(Name = "ERechnung_ExportPDF")] public static void ExportPdf() => Run(() => Commands.Export(ExportKind.ZugferdPdf));
    [ExcelCommand(Name = "ERechnung_Import")] public static void Import() => Run(Commands.Import);
    [ExcelCommand(Name = "ERechnung_Eingang")] public static void Inbox() => Run(Commands.ShowInbox);

    /// <summary>
    /// Switches to unattended mode for automated tests: no dialogs, confirmations answered with "yes",
    /// the next save dialog returns <paramref name="savePath"/>, the next open dialog returns the
    /// '|'-separated <paramref name="openPaths"/>. Empty <paramref name="savePath"/> and <paramref name="openPaths"/> ends test mode.
    /// </summary>
    [ExcelFunction(Name = "ERechnung_Testmodus", IsMacroType = true, IsHidden = true)]
    public static string TestMode(string savePath, string openPaths)
    {
        var on = !string.IsNullOrEmpty(savePath) || !string.IsNullOrEmpty(openPaths);
        Ui.Automation = on ? new Ui.AutomationSettings(savePath, openPaths.Split('|', StringSplitOptions.RemoveEmptyEntries)) : null;
        return on ? "Testmodus an" : "Testmodus aus";
    }

    /// <summary>Messages that would have been shown in dialogs during test mode (and clears them).</summary>
    [ExcelFunction(Name = "ERechnung_Testprotokoll", IsMacroType = true, IsHidden = true)]
    public static string TestLog()
    {
        if (Ui.Automation is null) return "";
        var text = string.Join("\n---\n", Ui.Automation.Log);
        Ui.Automation.Log.Clear();
        return text;
    }

    private static void Run(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            Ui.Error((ex is ERechnung.Core.InvoiceDataException ? "" : "Unerwarteter Fehler:\n\n") + ex.Message);
        }
    }
}
