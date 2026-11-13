using System.Windows.Forms;
using ExcelDna.Integration;

namespace ERechnung.ExcelAddIn;

/// <summary>Message boxes and file dialogs owned by the Excel window.</summary>
internal static class Ui
{
    private const string Caption = "E-Rechnung";

    private sealed class ExcelWindow : IWin32Window
    {
        public IntPtr Handle => ExcelDnaUtil.WindowHandle;
    }

    private static readonly IWin32Window Owner = new ExcelWindow();

    /// <summary>Unattended mode for automated tests (see <see cref="Macros.TestMode"/>); null = normal dialogs.</summary>
    public static AutomationSettings? Automation { get; set; }

    public sealed record AutomationSettings(string SavePath, string[] OpenPaths)
    {
        public List<string> Log { get; } = new();
    }

    public static void Info(string text) => Show(text, MessageBoxIcon.Information, "INFO");

    public static void Warn(string text) => Show(text, MessageBoxIcon.Warning, "WARN");

    public static void Error(string text) => Show(text, MessageBoxIcon.Error, "ERROR");

    private static void Show(string text, MessageBoxIcon icon, string tag)
    {
        if (Automation is { } a)
        {
            a.Log.Add($"{tag}: {text}");
            return;
        }
        MessageBox.Show(Owner, text, Caption, MessageBoxButtons.OK, icon);
    }

    public static bool Confirm(string text)
    {
        if (Automation is { } a)
        {
            a.Log.Add("CONFIRM(yes): " + text);
            return true;
        }
        return MessageBox.Show(Owner, text, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
    }

    public static string? SaveFile(string title, string filter, string defaultName, string? initialDir)
    {
        if (Automation is { } a) return string.IsNullOrEmpty(a.SavePath) ? null : a.SavePath;
        using var dlg = new SaveFileDialog
        {
            Title = title,
            Filter = filter,
            FileName = defaultName,
            OverwritePrompt = true,
            AddExtension = true,
        };
        if (!string.IsNullOrEmpty(initialDir) && Directory.Exists(initialDir)) dlg.InitialDirectory = initialDir;
        return dlg.ShowDialog(Owner) == DialogResult.OK ? dlg.FileName : null;
    }

    public static string[] OpenFiles(string title, string filter, string? initialDir)
    {
        if (Automation is { } a) return a.OpenPaths;
        using var dlg = new OpenFileDialog
        {
            Title = title,
            Filter = filter,
            Multiselect = true,
            CheckFileExists = true,
        };
        if (!string.IsNullOrEmpty(initialDir) && Directory.Exists(initialDir)) dlg.InitialDirectory = initialDir;
        return dlg.ShowDialog(Owner) == DialogResult.OK ? dlg.FileNames : Array.Empty<string>();
    }
}
