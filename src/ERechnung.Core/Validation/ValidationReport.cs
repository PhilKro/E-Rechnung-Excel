namespace ERechnung.Core.Validation;

public enum Severity
{
    Error,
    Warning,
    Info,
}

/// <summary>Where a message came from.</summary>
public enum MessageSource
{
    /// <summary>The add-in's own checks (<see cref="BuiltInValidator"/>), available offline and instantly.</summary>
    BuiltIn,

    /// <summary>The official KoSIT validator (XSD + EN 16931 + XRechnung Schematron).</summary>
    Kosit,
}

/// <param name="Code">Rule id, e.g. "BR-DE-15" or "BR-CO-10".</param>
/// <param name="Text">German message for the user.</param>
/// <param name="FieldKey">Sheet field the message refers to (see <see cref="Sheet.FieldKeys"/>), if known.</param>
/// <param name="Location">XPath (KoSIT) or line reference.</param>
public sealed record ValidationMessage(
    Severity Severity,
    string Code,
    string Text,
    MessageSource Source = MessageSource.BuiltIn,
    string? FieldKey = null,
    string? Location = null);

/// <summary>All messages of one validation run.</summary>
public sealed class ValidationReport
{
    public List<ValidationMessage> Messages { get; } = new();

    /// <summary>Result of the KoSIT validator: true = accepted, false = rejected, null = not run.</summary>
    public bool? KositAccepted { get; set; }

    /// <summary>Why KoSIT did not run (not installed, Java missing, …), if applicable.</summary>
    public string? KositSkippedReason { get; set; }

    /// <summary>Path of the HTML report written by KoSIT, if any.</summary>
    public string? KositHtmlReport { get; set; }

    /// <summary>Scenario KoSIT matched, e.g. "XRechnung UBL Invoice".</summary>
    public string? KositScenario { get; set; }

    public int ErrorCount => Messages.Count(m => m.Severity == Severity.Error);
    public int WarningCount => Messages.Count(m => m.Severity == Severity.Warning);

    /// <summary>Valid = no errors and KoSIT (if it ran) accepted the document.</summary>
    public bool IsValid => ErrorCount == 0 && KositAccepted != false;

    public void Add(Severity severity, string code, string text, string? fieldKey = null, string? location = null) =>
        Messages.Add(new ValidationMessage(severity, code, text, MessageSource.BuiltIn, fieldKey, location));

    /// <summary>One-line German summary for message boxes and the status bar.</summary>
    public string Summary
    {
        get
        {
            var kosit = KositAccepted switch
            {
                true => "KoSIT: akzeptiert",
                false => "KoSIT: abgelehnt",
                null => KositSkippedReason is null ? "KoSIT: nicht ausgeführt" : $"KoSIT: nicht ausgeführt ({KositSkippedReason})",
            };
            var verdict = IsValid ? "Gültig" : "Ungültig";
            return $"{verdict} – {ErrorCount} Fehler, {WarningCount} Warnungen. {kosit}.";
        }
    }
}
