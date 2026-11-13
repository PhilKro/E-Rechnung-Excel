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

Step-by-step instructions for everyday use: **[docs/USER-GUIDE.md](docs/USER-GUIDE.md)**.

## Example

A complete, valid invoice to try the add-in with. **E-Rechnung → Beispiel** fills in exactly this; to see how
entering an invoice works, type it into a sheet created with **Neue Rechnung**. The official KoSIT validator
accepts it as XRechnung UBL, XRechnung CII and ZUGFeRD PDF. All companies, numbers and the IBAN are fictitious.

**Rechnung**

| Field | Value |
|---|---|
| Rechnungsnummer | RE-2027-0042 |
| Rechnungsdatum | 15.03.2027 |
| Rechnungsart | 380 - Rechnung |
| Währung | EUR |
| Leistungsdatum | 10.03.2027 |
| Fälligkeitsdatum | 29.03.2027 |
| Zahlungsbedingungen | Zahlbar innerhalb 14 Tagen ohne Abzug. |
| Käuferreferenz / Leitweg-ID | 04011000-12345-34 |
| Bestellnummer | PO-7781 |
| Bemerkung | Vielen Dank für Ihren Auftrag. |

**Verkäufer** and **Käufer**

| Field | Verkäufer | Käufer |
|---|---|---|
| Name / Firma | Musterbau GmbH | Beispiel Handels AG |
| Kundennummer | | K-1001 |
| Straße, Hausnr. | Hauptstraße 1 | Marktplatz 5 |
| PLZ / Ort / Land | 10115 / Berlin / DE | 80331 / München / DE |
| USt-IdNr. | DE123456789 | DE987654321 |
| Steuernummer | 30/123/45678 | |
| Elektronische Adresse | rechnung@musterbau.example | eingangsrechnungen@beispiel-handel.example |
| Ansprechpartner | Erika Mustermann | Max Beispiel |
| Telefon | +49 30 1234567 | |
| E-Mail | buchhaltung@musterbau.example | |

Leave *Schema elektr. Adresse* empty; it is detected automatically.

**Zahlung**

| Field | Value |
|---|---|
| Zahlungsart | 58 - SEPA-Überweisung |
| IBAN | DE02120300000000202051 |
| BIC | BYLADEM1001 |
| Kontoinhaber | Musterbau GmbH |
| Verwendungszweck | RE-2027-0042 |

**Positionen** (the *Nettobetrag* column is a formula)

| Pos. | Bezeichnung | Beschreibung | Artikelnr. | Menge | Einheit | Einzelpreis netto | USt-Kategorie | USt-Satz % |
|---|---|---|---|---|---|---|---|---|
| 1 | Montagearbeiten | Montage Regalsystem | MON-01 | 7,5 | HUR | 68,00 | S | 19 |
| 2 | Regalsystem Typ 200 | | REG-200 | 3 | C62 | 249,90 | S | 19 |
| 3 | Fachbuch Lagerlogistik | | BUCH-1 | 2 | C62 | 39,95 | S | 7 |

**Prüfen** then shows *Gültig* and these totals:

| | |
|---|---|
| Summe Positionen netto | 1.339,60 |
| USt 19 % auf 1.259,70 | 239,34 |
| USt 7 % auf 79,90 | 5,59 |
| **Zahlbetrag** | **1.584,53** |

Things to try next:

- Clear *Käuferreferenz / Leitweg-ID* and click **Prüfen**: the field turns red and the **Prüfbericht** lists
  `BR-DE-15`.
- Make it a Kleinunternehmer invoice:
  - delete the seller's USt-IdNr. and set *Verkäufer-Kennung* to `LIEF-4711`;
  - set every line to USt-Kategorie `E` with rate `0`;
  - set *Grund der Steuerbefreiung* to "Kein Ausweis von Umsatzsteuer, da Kleinunternehmer gemäß § 19 UStG.".
- Export a **ZUGFeRD-PDF**, then load it with **E-Rechnung öffnen**. It appears in **Rechnungseingang** as
  *gültig*.

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
