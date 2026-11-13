# E-Rechnung für Excel

An Excel add-in for creating, sending, receiving and checking German e-invoices
(**XRechnung 3.0** and **ZUGFeRD 2.3 / Factur-X**), as required by §14 UStG from 2025/2027/2028.
The user interface is German; this README and the code are in English.

| Stage | What it does |
|---|---|
| Create | Ribbon tab **E-Rechnung → Neue Rechnung** adds an invoice sheet with every field an XRechnung needs, drop-downs for codes and a line table. **Beispiel** adds a filled-in example. |
| Send | Export as **XRechnung (UBL)**, **XRechnung (CII)** or **ZUGFeRD-PDF** (the sheet as PDF with the XML embedded). |
| Receive | **E-Rechnung öffnen** reads XRechnung XML and ZUGFeRD/Factur-X PDFs (several at once), shows a single invoice in its own sheet and logs every file in **Rechnungseingang** (amounts, IBAN, due date, validation result). |
| Check | **Prüfen** computes the totals, checks the main EN 16931 / XRechnung rules, marks wrong fields and writes a **Prüfbericht** with links to the fields. With the official **KoSIT validator** installed, every check, export and import is also validated exactly like the public-sector invoice portals do. |

## Installation (users)

1. Install the [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (x64 for 64-bit Excel).
2. Copy `ERechnung-AddIn64-packed.xll` (64-bit Excel) or `ERechnung-AddIn-packed.xll` (32-bit) to a local folder.
3. Excel → *Datei → Optionen → Add-Ins → Verwalten: Excel-Add-Ins → Gehe zu → Durchsuchen* → select the `.xll`.
4. Optional but recommended: tab **E-Rechnung → Status / Installieren** installs the KoSIT validator
   (~60 MB, per user, no admin rights). The checkbox **KoSIT-Prüfung** turns it on or off.

## Building (developers)

Requires Windows, the .NET 8 SDK and – for the smoke test – desktop Excel.

```powershell
./build.ps1                    # build + tests
./tools/setup-kosit.ps1        # official validator for tests and add-in
./tests/excel-smoke.ps1        # end-to-end test in a hidden Excel instance
```

The add-in is written to `src/ERechnung.ExcelAddIn/bin/Debug/net8.0-windows/publish/`.
For F5 debugging, start Excel with `bin/Debug/net8.0-windows/ERechnung-AddIn64.xll`.

See [AGENTS.md](AGENTS.md) for architecture, conventions, library quirks and known limitations.

## Scope and limitations

Covers typical invoices: standard and reduced VAT, §19 UStG (Kleinunternehmer), reverse charge (§13b),
intra-community supply, export, credit notes and corrections, SEPA credit transfer and direct debit, Leitweg-IDs
for public-sector recipients. Not yet supported: document-level discounts/surcharges, delivery addresses,
attachments, card payments. This tool does not replace tax advice; the KoSIT validator checks the format and
the business rules, not whether the invoice is correct for your tax situation.

Built on [ZUGFeRD-csharp](https://github.com/stephanstapel/ZUGFeRD-csharp), [Excel-DNA](https://excel-dna.net)
and the [KoSIT validator](https://github.com/itplr-kosit/validator).
