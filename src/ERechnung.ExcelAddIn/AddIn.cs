using ExcelDna.Integration;

namespace ERechnung.ExcelAddIn;

/// <summary>Excel-DNA entry point. The ribbon (<see cref="Ribbon"/>) is discovered automatically.</summary>
public sealed class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        // Unhandled exceptions in ribbon callbacks would otherwise disappear silently.
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "Fehler: " + ex);
    }

    public void AutoClose()
    {
    }
}
