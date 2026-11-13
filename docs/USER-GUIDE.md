# User guide

How to create, send, receive and check e-invoices with the **E-Rechnung** add-in for Excel.
The add-in itself is in German; button and field names below are quoted exactly as they appear in Excel.

## 1. Install

1. Install the [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (x64 for 64-bit Excel).
2. Copy `ERechnung-AddIn64-packed.xll` (64-bit Excel) or `ERechnung-AddIn-packed.xll` (32-bit Excel) to a local
   folder, e.g. `Documents\E-Rechnung`. Not sure which Excel you have? *Datei → Konto → Info zu Excel* shows
   "64-Bit" or "32-Bit".
3. Excel → *Datei → Optionen → Add-Ins* → at the bottom *Verwalten: Excel-Add-Ins* → *Gehe zu…* →
   *Durchsuchen…* → select the `.xll` → OK. A new ribbon tab **E-Rechnung** appears.
   The add-in is not signed, so Excel may ask whether to enable it – choose *Dieses Add-In nur für diese Sitzung
   aktivieren* or *Aktivieren*.
4. Recommended: **E-Rechnung → Status / Installieren** installs the official KoSIT validator (~60 MB, per user,
   no admin rights, needs internet once). Wait until the PowerShell window says `KoSIT validator installed`.
   The checkbox **KoSIT-Prüfung** turns the official check on or off.

## 2. The ribbon

| Group | Button | What it does |
|---|---|---|
| Rechnung erstellen | **Neue Rechnung** | Adds an empty invoice sheet. |
| | **Beispiel** | Adds a filled-in, valid example invoice (see the [README](../README.md#example)). |
| | **Prüfen** | Computes the totals, checks the invoice and writes the sheet **Prüfbericht**. |
| Senden (Export) | **XRechnung (UBL)** | Saves the active invoice sheet as XRechnung XML (UBL syntax). |
| | **XRechnung (CII)** | Same, CII syntax. |
| | **ZUGFeRD-PDF** | Saves the sheet as PDF with the invoice XML embedded (ZUGFeRD 2.3 / Factur-X). |
| Empfangen (Import) | **E-Rechnung öffnen** | Reads one or more received XRechnung/ZUGFeRD files. |
| | **Rechnungseingang** | Shows the log of all received invoices. |
| Offizielle Prüfung | **KoSIT-Prüfung** | Turns the official validator on/off. |
| | **Status / Installieren** | Shows the installed version or installs the validator. |

## 3. Create an invoice

1. **Neue Rechnung** adds a sheet. Top left are the header fields, at the bottom the line table, on the right
   the totals.
   - **Bold** labels are mandatory for XRechnung.
   - Column C holds hints.
   - The BT number (e.g. BT-10) is the field's number in the EN 16931 standard.
2. Fill in the fields. Fields with a drop-down (Rechnungsart, Währung, Land, Zahlungsart, Einheit, USt-Kategorie)
   accept the list entry or just the code (`380`, `EUR`, `DE`, `58`, `HUR`, `S`). Dates and numbers may be typed
   the German way (`15.03.2027`, `1.234,50`).
3. Enter one row per line item. **Nettobetrag** is a formula (Menge × Einzelpreis, rounded) – do not overwrite
   it. Rows where Bezeichnung, Beschreibung, Artikelnr., Menge and Einzelpreis are all empty are ignored; to add
   rows, type directly below the table.
4. Click **Prüfen**. The totals appear on the right, wrong or missing fields turn red, and the sheet
   **Prüfbericht** lists every finding. Each finding has:
   - a rule ID (e.g. `BR-DE-15`);
   - a link back to the field;
   - a link to the official KoSIT report, if the validator is installed.

   *Fehler* (errors) make the invoice invalid. *Warnungen* are recommendations.

Saving the workbook keeps the invoice; reopening it later works as long as the add-in is loaded.

### Fields that are easy to get wrong

| Field | What to enter |
|---|---|
| Käuferreferenz / Leitweg-ID (BT-10) | Always mandatory in XRechnung. Public-sector customers give you a **Leitweg-ID** (e.g. `04011000-12345-34`); for companies use the reference they ask for, e.g. their order number. |
| Elektronische Adresse (BT-34 / BT-49) | Usually the e-mail address for invoices. Leave *Schema* empty – it is detected automatically (e-mail → `EM`, Leitweg-ID → `0204`). |
| USt-IdNr. / Steuernummer / Verkäufer-Kennung | At least one of USt-IdNr. or Steuernummer. Without a USt-IdNr. (e.g. Kleinunternehmer) also fill **Verkäufer-Kennung**, e.g. your supplier number at the customer. |
| Ansprechpartner, Telefon, E-Mail (seller) | All three are mandatory in XRechnung. |
| Leistungsdatum or Leistungszeitraum | One of them is required by §14 UStG. |
| Fälligkeitsdatum / Zahlungsbedingungen | One of them is required. |
| IBAN | Required for *58 - SEPA-Überweisung*. The check digits are verified. |

### Special VAT cases

| Case | USt-Kategorie | USt-Satz % | Also fill |
|---|---|---|---|
| Normal 19 % / reduced 7 % | `S` | 19 / 7 | – |
| Kleinunternehmer (§19 UStG) | `E` | 0 | Grund der Steuerbefreiung, e.g. "Kein Ausweis von Umsatzsteuer, da Kleinunternehmer gemäß § 19 UStG."; Verkäufer-Kennung if you have no USt-IdNr. |
| Reverse charge (§13b UStG) | `AE` | 0 | Grund der Steuerbefreiung ("Steuerschuldnerschaft des Leistungsempfängers"), Befreiungsgrund-Code `VATEX-EU-AE`, USt-IdNr. of the buyer |
| Intra-community supply | `K` | 0 | Grund der Steuerbefreiung, USt-IdNr. of seller and buyer, Leistungsdatum |
| Export to non-EU country | `G` | 0 | Grund der Steuerbefreiung |

### Credit notes and corrections

Choose Rechnungsart **381 - Gutschrift** or **384 - Rechnungskorrektur**, enter positive amounts, and fill
**Bezug: Rechnungsnummer** and **Bezug: Rechnungsdatum** with the original invoice.

### Direct debit

Zahlungsart **59 - SEPA-Lastschrift** needs **Mandatsreferenz**, **Gläubiger-ID** and **IBAN des Zahlers**.

## 4. Send an invoice

1. Select the invoice sheet.
2. Click the format your customer wants:
   - **XRechnung (UBL)** – the usual choice for public-sector customers and their invoice portals (ZRE, OZG-RE,
     Peppol).
   - **XRechnung (CII)** – the same content in the other syntax; some customers prefer it.
   - **ZUGFeRD-PDF** – a readable PDF with the XML inside; popular with companies because people can open it
     like any PDF.
3. The add-in checks the invoice first. If it finds errors it asks whether to save anyway – normally fix the
   errors instead. With KoSIT active, the exported file is checked again by the official validator.
4. Choose where to save and send the file to your customer (e-mail, portal upload, Peppol provider).

A plain PDF printed from Excel is **not** an e-invoice. Only the XML files and the ZUGFeRD PDF are.

## 5. Receive invoices

1. **E-Rechnung öffnen** → select one or more `.xml` or `.pdf` files (Ctrl+click for several).
2. Every file gets a row in the sheet **Rechnungseingang**:
   - seller, invoice number and date;
   - net, VAT, gross and amount due;
   - due date, IBAN and payment reference;
   - the check result (*gültig* / *ungültig*).

   Files that are not e-invoices (e.g. a plain PDF) are listed as *nicht lesbar*.
3. If you open a single file, the invoice is also shown in its own sheet **Eingang \<Rechnungsnr.\>**. Amounts
   there are the values stated in the received invoice. They are not recomputed, so arithmetic mistakes of the
   sender show up as findings.
4. Before paying, check that the IBAN and amounts match what you expect – a valid e-invoice can still be a
   fraudulent one.

Keep the original files: §14b UStG requires archiving the received XML/PDF unchanged for 8 years. The Excel
sheet is only a view.

## 6. Troubleshooting

| Problem | Solution |
|---|---|
| No **E-Rechnung** tab | The add-in is not loaded. Check *Datei → Optionen → Add-Ins*. Check that the .NET 8 Desktop Runtime is installed and that the 32/64-bit version matches your Excel. |
| "Das aktive Blatt ist kein Rechnungsblatt" | Select a sheet created with **Neue Rechnung** or **Beispiel**. |
| Excel freezes for ~10 seconds | That is the KoSIT validator running. Untick **KoSIT-Prüfung** for quick checks and turn it on before sending. |
| KoSIT reports an error the built-in check did not | KoSIT is the authority. The finding in the **Prüfbericht** links to the full HTML report. |
| "keine eingebettete Rechnungs-XML" | The PDF is a normal PDF, not ZUGFeRD. Ask the sender for an e-invoice. |
| Validator outdated | Click **Status / Installieren** again, or run `tools/setup-kosit.ps1`. The XRechnung rules are updated about twice a year. |

This tool helps you produce correctly formatted invoices. It does not replace tax advice: the validator
checks format and business rules, not whether the invoice is correct for your tax situation.
