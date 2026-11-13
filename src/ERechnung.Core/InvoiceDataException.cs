namespace ERechnung.Core;

/// <summary>
/// Thrown when invoice data cannot be converted (unknown code, unreadable file, …).
/// The message is German and meant to be shown to the Excel user as-is.
/// </summary>
public sealed class InvoiceDataException : Exception
{
    public InvoiceDataException(string message) : base(message) { }
    public InvoiceDataException(string message, Exception inner) : base(message, inner) { }
}
