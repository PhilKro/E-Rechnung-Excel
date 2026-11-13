# AGENTS.md

Guidance for AI coding agents (and humans) working on this repository.

## What this is

An Excel add-in (Windows desktop Excel, Excel-DNA, .NET 8) for German e-invoices according to EN 16931:

1. **Create** invoices in a worksheet (one sheet = one invoice).
2. **Send / export** as XRechnung 3.0 (UBL or CII) or ZUGFeRD 2.3 / Factur-X PDF (profile EN 16931).
3. **Receive / import** XRechnung XML and ZUGFeRD PDFs into a sheet and a log ("Rechnungseingang").
4. **Validate** with built-in rules and the official KoSIT validator (XSD + EN 16931 + XRechnung Schematron).
5. **Excel UI**: ribbon tab "E-Rechnung", invoice template, report sheet, inbox sheet.

Legal background (Wachstumschancengesetz, §14 UStG): since 2025-01-01 every German business must be able to
*receive* e-invoices; from 2027-01-01 businesses with > €800k turnover must *send* them; from 2028-01-01 all
domestic B2B invoices must be e-invoices. A plain PDF is **not** an e-invoice; ZUGFeRD (PDF + XML) is.

## Repository layout

```
src/ERechnung.Core/          net8.0, no Excel dependency – everything testable lives here
  Model/Invoice.cs           neutral EN 16931 model (BT-/BG- numbers in doc comments)
  Codes/CodeLists.cs         code lists with German labels for drop-downs
  Calculation/               totals + VAT breakdown (EN 16931 rounding)
  Zugferd/                   mapping Invoice <-> ZUGFeRD-csharp InvoiceDescriptor, enum<->code helper
  IO/                        InvoiceExporter (XRechnung/ZUGFeRD), InvoiceImporter, PDF attachment extraction
  Validation/                BuiltInValidator, KositValidator (external Java process), ValidationReport
  Sheet/                     SheetLayout (which fields exist, labels, keys) + SheetMapper (cell values <-> Invoice)
  Samples/DemoInvoices.cs    fictitious invoices used by the "Beispiel" button and the tests
src/ERechnung.ExcelAddIn/    net8.0-windows, Excel-DNA; thin layer over Core
  Ribbon.cs                  ribbon XML + callbacks (dispatch only)
  Commands.cs                the actual actions (create, check, export, import, KoSIT)
  InvoiceSheet.cs            renders SheetLayout, reads/writes named cells and the line table
  ReportSheet.cs / InboxSheet.cs / ExcelHelpers.cs / Ui.cs / Settings.cs
  Macros.cs                  ERechnung_* macros (Application.Run) + unattended test mode
tests/ERechnung.Core.Tests/  xUnit: round trips, sheet mapping, rules, KoSIT acceptance
tests/excel-smoke.ps1        end-to-end test that drives a hidden Excel through all stages
tools/setup-kosit.ps1        installs KoSIT validator + XRechnung config + portable JRE per user
```

## Build and test

The `dotnet` on PATH may have no SDK. `build.ps1` finds one (PATH, then the per-user install at
`%LOCALAPPDATA%\Microsoft\dotnet-sdk`). To install an SDK without admin rights:
`dotnet-install.ps1 -Channel 8.0 -InstallDir $env:LOCALAPPDATA\Microsoft\dotnet-sdk -NoPath`.

```powershell
./build.ps1                  # build + all tests (Debug)
./build.ps1 -Release         # Release build
./tools/setup-kosit.ps1      # once per machine: enables the KoSIT tests and the official check in Excel
./tests/excel-smoke.ps1      # after add-in changes: needs desktop Excel 64-bit, runs ~2 min
dotnet test tests/ERechnung.Core.Tests --filter "FullyQualifiedName!~KositTests"   # fast loop (~5 s)
```

- The very first restore/build on a fresh machine can take ~5 minutes (antivirus scanning); later builds take seconds.
- KoSIT tests start one JVM per test (~10 s each, ~3.5 min total). Without KoSIT installed they pass trivially and
  print "SKIPPED" – that is **not** a successful compliance check.
- Add-in output: `src/ERechnung.ExcelAddIn/bin/<Config>/net8.0-windows/publish/ERechnung-AddIn64-packed.xll`
  (single file, all dependencies packed; `ERechnung-AddIn-packed.xll` for 32-bit Excel).
  Users need the .NET 8 Desktop Runtime.

## Rules for changes

**Compliance is verified, not assumed.** The KoSIT validator is the authority. Any change to mapping, export or
the model must keep `KositTests` green with KoSIT installed. If you add a feature (new field, new VAT case,
new payment means), add a sample to `KositTests.Sample()` and make KoSIT accept it. Do not "fix" a KoSIT finding
by weakening a test.

**Core stays Excel-free.** Anything that can be unit tested goes into `ERechnung.Core`. The add-in only moves
values between cells and Core, and shows results. Never reference Excel interop from Core.

**One model.** `Invoice` is shared by sheet, export, import and validation. Totals for outgoing invoices are
always computed by `InvoiceCalculator` – never read from the sheet. Received invoices keep their stated totals
(`StatedTotals`, `StatedNetAmount`) so inconsistencies stay visible.

**Adding a header field** (checklist):
1. Property on `Invoice`/`Party`/`PaymentInstructions` with the BT number in the doc comment.
2. Key in `FieldKeys`, `FieldDef` in `SheetLayout.Sections` (German label, BT number, hint).
3. Read and write in `SheetMapper`.
4. Map both directions in `DescriptorMapper`.
5. Rules in `BuiltInValidator` if mandatory; a test; KoSIT sample if it changes the XML.
Existing workbooks simply lack the new named cell – `InvoiceSheet.Read` treats missing names as empty. Never
rename a key: Excel names (`ER_<key>`) are derived from it and old workbooks would break. Labels may change.
Line table columns are matched by header text, so renaming a header breaks old sheets too.

**Rule IDs must match the official ones.** Built-in messages use the EN 16931 / XRechnung rule IDs (BR-*,
BR-DE-*, PEPPOL-EN16931-R*) so they line up with KoSIT reports. Check the official text in the installed
configuration: `%LOCALAPPDATA%\ERechnungExcel\kosit\config\resources\xrechnung\<ver>\xsl\XRechnung-UBL-validation.xsl`.
Use `Severity.Warning` for "soll" rules (e.g. BR-DE-19, BR-DE-26, BR-DE-27) and `Error` for "muss" rules. The
built-in check must not report errors for documents KoSIT accepts (`BuiltInValidator_HasNoErrors_ForValidSamples`).

**Language.** Code, comments, commit messages, docs: English. Everything the Excel user sees (labels, messages,
`InvoiceDataException` texts, ribbon): German.

**PowerShell scripts must be ASCII-only.** Windows PowerShell 5.1 reads BOM-less UTF-8 as ANSI; umlauts in a
script break string comparisons silently. Match umlauts with `.` in regexes or build them with `[char]0xFC`.

**Excel automation.** Never stop Excel processes broadly – the user may have workbooks open. The smoke test
records the PID of the instance it starts and ends only that one. It copies the `.xll` to its output folder so a
lingering instance never locks the build output.

## Library notes (ZUGFeRD-csharp 18.0.1)

- Enum values are not the codes (`InvoiceType.Invoice` = 22, code "380"). Always convert via `Zugferd/EnumCodes`,
  which reads `EnumStringValueAttribute`; the library's own helpers are internal.
- `ReferenceOrderNo` is BT-10 (buyer reference / Leitweg-ID); `OrderNo` is BT-13.
- `BusinessProcess` (BT-23) must be set, otherwise CII fails PEPPOL-EN16931-R001.
- An empty payment-terms text produces an empty CII element (PEPPOL-EN16931-R008), so `DescriptorMapper` derives
  a text from the due date.
- UBL **CreditNote** output carries no due date (BT-9) – known library limitation, covered by a test.
- `InvoicePdfProcessor` (ZUGFeRD.PDF-csharp, PDFsharp-extended) embeds the XML into an existing PDF; the visual
  PDF comes from Excel's `ExportAsFixedFormat`. Use `PdfAttachments` to get the raw embedded XML back.

## KoSIT validator notes

- Installed per user by `tools/setup-kosit.ps1` (latest releases from GitHub `itplr-kosit/validator` and
  `itplr-kosit/validator-configuration-xrechnung`, Temurin JRE from adoptium.net). Re-run it to update; the
  XRechnung configuration is released roughly twice a year and must be kept current.
- It probes stdin to detect piped input and crashes with "Incorrect function" when stdin is NUL. `KositValidator`
  therefore always redirects stdin. Run it manually from a real console, not from a non-interactive shell.
- The process exit code reflects the validation result; `KositValidator` relies on the generated
  `*-report.xml` (namespace `http://www.xoev.de/de/validator/varl/1`) instead.
- Batch several documents into one `ValidateMany` call – JVM start and Schematron compilation dominate.

## Known limitations (candidates for future work)

- No document-level allowances/charges (BG-20/21), no line-level allowances, no delivery address (BG-13/15),
  no attachments (BG-24), no payee/tax representative, no card payments (needs BG-18).
- One VAT exemption reason per invoice (applied to all exempt categories).
- The ZUGFeRD PDF's PDF/A-3 conformance has not been checked with veraPDF; Excel's PDF export is not PDF/A by
  default. Verify before relying on it for archiving.
- KoSIT runs synchronously on Excel's main thread (UI freezes ~10 s per check).
- The visual PDF is the worksheet as printed (hint column hidden), not a designed invoice layout.
- Not signed: Excel may show a security notice for the unsigned `.xll`.
