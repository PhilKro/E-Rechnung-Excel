using System.Text.Json;
using ERechnung.Core.Validation;

namespace ERechnung.ExcelAddIn;

/// <summary>Per-user settings, stored in %APPDATA%\ERechnungExcel\settings.json.</summary>
public sealed class Settings
{
    /// <summary>Run the official KoSIT validator on check, export and import (if installed).</summary>
    public bool UseKosit { get; set; } = true;

    /// <summary>KoSIT installation folder; empty = default (%LOCALAPPDATA%\ERechnungExcel\kosit).</summary>
    public string? KositDir { get; set; }

    private static string FilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ERechnungExcel", "settings.json");

    private static Settings? _current;

    public static Settings Current => _current ??= Load();

    private static Settings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new Settings();
        }
        catch (Exception)
        {
            // Corrupt settings must never block Excel; fall back to defaults.
        }
        return new Settings();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }

    /// <summary>The KoSIT validator if enabled and installed; otherwise null and a German reason.</summary>
    public KositValidator? TryGetKosit(out string? reason)
    {
        if (!UseKosit)
        {
            reason = "abgeschaltet";
            return null;
        }
        return KositValidator.TryCreateDefault(out reason, KositDir);
    }
}
