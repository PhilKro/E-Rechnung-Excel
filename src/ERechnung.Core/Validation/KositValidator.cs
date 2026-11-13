using System.Diagnostics;
using System.Text;
using System.Xml;

namespace ERechnung.Core.Validation;

/// <summary>Where the KoSIT validator, its configuration and Java are located.</summary>
public sealed class KositSettings
{
    /// <summary>Default install folder used by <c>tools/setup-kosit.ps1</c>.</summary>
    public static string DefaultInstallDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ERechnungExcel", "kosit");

    public string InstallDir { get; init; } = DefaultInstallDir;

    /// <summary>java.exe: the portable JRE in the install folder if present, otherwise "java" from PATH.</summary>
    public string JavaPath { get; init; } = "";

    public string ValidatorJar => Path.Combine(InstallDir, "validator.jar");
    public string ScenariosFile => Path.Combine(InstallDir, "config", "scenarios.xml");
    public string RepositoryDir => Path.Combine(InstallDir, "config");

    /// <summary>Time limit for one validator run.</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromMinutes(3);

    public static KositSettings FromInstallDir(string? installDir = null)
    {
        var dir = string.IsNullOrWhiteSpace(installDir) ? DefaultInstallDir : installDir;
        var bundledJava = Path.Combine(dir, "jre", "bin", "java.exe");
        return new KositSettings { InstallDir = dir, JavaPath = File.Exists(bundledJava) ? bundledJava : "java" };
    }

    /// <summary>German description of what is missing, or <c>null</c> if the installation looks complete.</summary>
    public string? FindProblem()
    {
        if (!File.Exists(ValidatorJar)) return $"validator.jar nicht gefunden in {InstallDir}";
        if (!File.Exists(ScenariosFile)) return $"XRechnung-Konfiguration (scenarios.xml) nicht gefunden in {RepositoryDir}";
        if (JavaPath != "java" && !File.Exists(JavaPath)) return $"Java nicht gefunden: {JavaPath}";
        return null;
    }
}

/// <summary>Result of validating one document with KoSIT.</summary>
public sealed class KositResult
{
    public required string DocumentName { get; init; }
    public bool Accepted { get; init; }
    public string? Scenario { get; init; }
    public List<ValidationMessage> Messages { get; } = new();

    /// <summary>The validator's HTML report (human readable, as published by KoSIT), copied to a persistent folder.</summary>
    public string? HtmlReport { get; set; }
}

/// <summary>
/// Runs the official KoSIT validator (https://github.com/itplr-kosit/validator) with the XRechnung configuration
/// as an external Java process and parses its XML report. It checks XSD + EN 16931 Schematron + XRechnung Schematron,
/// i.e. exactly what public-sector invoice portals check.
/// </summary>
public sealed class KositValidator
{
    private readonly KositSettings _settings;

    public KositValidator(KositSettings settings) => _settings = settings;

    public KositSettings Settings => _settings;

    /// <summary>Folder where HTML reports are kept so the user can open them later.</summary>
    public static string ReportDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ERechnungExcel", "reports");

    /// <summary>Creates a validator for the default installation, or returns <c>null</c> with a German reason.</summary>
    public static KositValidator? TryCreateDefault(out string? reason, string? installDir = null)
    {
        var settings = KositSettings.FromInstallDir(installDir);
        reason = settings.FindProblem();
        return reason is null ? new KositValidator(settings) : null;
    }

    public KositResult Validate(byte[] xml, string documentName) =>
        ValidateMany(new[] { (documentName, xml) })[0];

    /// <summary>Validates several documents in one Java process (JVM start + Schematron compilation dominate the run time).</summary>
    public IReadOnlyList<KositResult> ValidateMany(IReadOnlyList<(string Name, byte[] Xml)> documents)
    {
        if (documents.Count == 0) return Array.Empty<KositResult>();

        var work = Path.Combine(Path.GetTempPath(), "ERechnungExcel", "kosit-" + Guid.NewGuid().ToString("N"));
        var inDir = Path.Combine(work, "in");
        var outDir = Path.Combine(work, "out");
        Directory.CreateDirectory(inDir);
        Directory.CreateDirectory(outDir);
        try
        {
            var inputs = new List<(string Name, string Stem, string Path)>();
            for (var i = 0; i < documents.Count; i++)
            {
                var stem = $"{i + 1:D3}_{SafeFileStem(documents[i].Name)}";
                var path = Path.Combine(inDir, stem + ".xml");
                File.WriteAllBytes(path, documents[i].Xml);
                inputs.Add((documents[i].Name, stem, path));
            }

            RunJava(outDir, inputs.Select(x => x.Path));

            var results = new List<KositResult>();
            foreach (var (name, stem, _) in inputs)
            {
                var reportXml = Path.Combine(outDir, stem + "-report.xml");
                if (!File.Exists(reportXml))
                    throw new InvalidOperationException($"KoSIT hat keinen Prüfbericht für \"{name}\" erzeugt.");
                var result = ParseReport(File.ReadAllBytes(reportXml), name);

                var html = Directory.EnumerateFiles(outDir, stem + "*.html").FirstOrDefault();
                if (html is not null)
                {
                    Directory.CreateDirectory(ReportDir);
                    var target = Path.Combine(ReportDir, $"{SafeFileStem(name)}_{DateTime.Now:yyyyMMdd-HHmmss}.html");
                    File.Copy(html, target, overwrite: true);
                    result.HtmlReport = target;
                }
                results.Add(result);
            }
            return results;
        }
        finally
        {
            try { Directory.Delete(work, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }

    private void RunJava(string outDir, IEnumerable<string> files)
    {
        var psi = new ProcessStartInfo(_settings.JavaPath)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            // The validator probes stdin to detect piped input; a NUL/console stdin makes that probe throw
            // ("Incorrect function"), so give it an empty pipe instead.
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        foreach (var a in new[] { "-Dfile.encoding=UTF-8", "-jar", _settings.ValidatorJar,
                     "-s", _settings.ScenariosFile, "-r", _settings.RepositoryDir, "-o", outDir, "-h" })
            psi.ArgumentList.Add(a);
        foreach (var f in files) psi.ArgumentList.Add(f);

        using var p = new Process { StartInfo = psi };
        var output = new StringBuilder();
        p.OutputDataReceived += (_, e) => { if (e.Data is not null) lock (output) output.AppendLine(e.Data); };
        p.ErrorDataReceived += (_, e) => { if (e.Data is not null) lock (output) output.AppendLine(e.Data); };
        try
        {
            p.Start();
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            throw new InvalidOperationException(
                $"Java konnte nicht gestartet werden ({_settings.JavaPath}). Bitte tools\\setup-kosit.ps1 ausführen. {ex.Message}", ex);
        }
        p.StandardInput.Close();
        p.BeginOutputReadLine();
        p.BeginErrorReadLine();
        if (!p.WaitForExit((int)_settings.Timeout.TotalMilliseconds))
        {
            try { p.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
            throw new TimeoutException($"Der KoSIT-Validator hat nicht innerhalb von {_settings.Timeout.TotalSeconds:0} s geantwortet.");
        }
        p.WaitForExit();   // flush async output

        // The validator's exit code reflects the validation outcome (e.g. rejected documents), not only
        // technical failures, so the reports decide. Only a run without any report is an error.
        if (!Directory.EnumerateFiles(outDir, "*-report.xml").Any())
        {
            string log;
            lock (output) log = output.ToString();
            throw new InvalidOperationException($"Der KoSIT-Validator ist fehlgeschlagen (Exit-Code {p.ExitCode}):{Environment.NewLine}{Tail(log, 1500)}");
        }
    }

    /// <summary>
    /// Parses a KoSIT report (namespace http://www.xoev.de/de/validator/varl/1). Relevant parts:
    /// <c>rep:scenarioMatched/rep:scenario/s:name</c>, every <c>rep:message</c> (attributes level, code, xpathLocation)
    /// and the final <c>rep:assessment</c> with <c>rep:accept</c> or <c>rep:reject</c>.
    /// </summary>
    internal static KositResult ParseReport(byte[] reportXml, string documentName)
    {
        var doc = new XmlDocument { XmlResolver = null };
        using (var ms = new MemoryStream(reportXml))
        using (var reader = XmlReader.Create(ms, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            doc.Load(reader);

        var all = doc.GetElementsByTagName("*").OfType<XmlElement>().ToList();
        var assessment = all.FirstOrDefault(e => e.LocalName == "assessment");
        var accepted = assessment?.ChildNodes.OfType<XmlElement>().Any(e => e.LocalName == "accept") == true;

        var scenarioMatched = all.FirstOrDefault(e => e.LocalName == "scenarioMatched");
        var scenarioName = scenarioMatched?.GetElementsByTagName("*").OfType<XmlElement>()
            .FirstOrDefault(e => e.LocalName == "scenario")?.ChildNodes.OfType<XmlElement>()
            .FirstOrDefault(e => e.LocalName == "name")?.InnerText.Trim();
        if (scenarioMatched is null && all.Any(e => e.LocalName == "noScenarioMatched"))
            scenarioName = "(kein passendes Szenario)";

        var result = new KositResult { DocumentName = documentName, Accepted = accepted, Scenario = scenarioName };

        foreach (var m in all.Where(e => e.LocalName == "message"))
        {
            var level = m.GetAttribute("level");
            var severity = level switch
            {
                "error" => Severity.Error,
                "warning" => Severity.Warning,
                _ => Severity.Info,
            };
            var code = m.GetAttribute("code");
            var text = m.InnerText.Trim();
            // Schematron messages look like "[BR-DE-15] Das Element ..." – keep the text without the duplicated id
            if (text.StartsWith('[') && text.IndexOf(']') is var close and > 0 && text[1..close] == code)
                text = text[(close + 1)..].Trim();
            result.Messages.Add(new ValidationMessage(severity, string.IsNullOrEmpty(code) ? "KoSIT" : code, text,
                MessageSource.Kosit, FieldKey: null, Location: NullIfEmpty(m.GetAttribute("xpathLocation"))));
        }

        if (scenarioMatched is null)
            result.Messages.Add(new ValidationMessage(Severity.Error, "KoSIT",
                "Kein Prüfszenario passt zu diesem Dokument (weder XRechnung noch EN 16931).", MessageSource.Kosit));

        return result;
    }

    private static string? NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s;

    internal static string SafeFileStem(string name) => SafeFileName(Path.GetFileNameWithoutExtension(name));

    /// <summary>Replaces characters that are invalid (or awkward) in file names, e.g. for invoice numbers with "/".</summary>
    public static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars().Concat(new[] { ' ', '#', '%', '&' }).ToHashSet();
        var safe = new string(name.Trim().Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        if (safe.Length > 60) safe = safe[..60];
        return safe.Length == 0 ? "rechnung" : safe;
    }

    private static string Tail(string s, int max) => s.Length <= max ? s : "…" + s[^max..];
}
