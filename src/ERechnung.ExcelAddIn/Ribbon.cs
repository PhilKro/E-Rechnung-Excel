using System.Runtime.InteropServices;
using ERechnung.Core;
using ExcelDna.Integration.CustomUI;

namespace ERechnung.ExcelAddIn;

/// <summary>The "E-Rechnung" ribbon tab. Callbacks only dispatch to <see cref="Commands"/> and catch errors.</summary>
[ComVisible(true)]
public sealed class Ribbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) => """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
          <ribbon>
            <tabs>
              <tab id="erTab" label="E-Rechnung">
                <group id="erCreate" label="Rechnung erstellen">
                  <button id="erNew" label="Neue Rechnung" size="large" imageMso="FileNew"
                          screentip="Neue Rechnung" supertip="Legt ein neues Rechnungsblatt mit allen Feldern einer XRechnung an."
                          onAction="OnNew" />
                  <button id="erDemo" label="Beispiel" imageMso="TableInsert"
                          screentip="Beispielrechnung" supertip="Legt ein ausgefülltes Beispiel an, das als gültige XRechnung exportiert werden kann."
                          onAction="OnDemo" />
                  <button id="erCheck" label="Prüfen" size="large" imageMso="AcceptInvitation"
                          screentip="Rechnung prüfen" supertip="Prüft das aktive Rechnungsblatt (eigene Prüfung und – falls installiert – offizieller KoSIT-Validator), berechnet die Summen und markiert fehlerhafte Felder."
                          onAction="OnCheck" />
                </group>
                <group id="erExport" label="Senden (Export)">
                  <button id="erUbl" label="XRechnung (UBL)" size="large" imageMso="XmlExport"
                          screentip="XRechnung UBL" supertip="Speichert die Rechnung als XRechnung 3.0 im UBL-Format – der übliche Standard für Behörden (Pflicht für B2G)."
                          onAction="OnExportUbl" />
                  <button id="erCii" label="XRechnung (CII)" imageMso="XmlExport"
                          screentip="XRechnung CII" supertip="Speichert die Rechnung als XRechnung 3.0 im CII-Format."
                          onAction="OnExportCii" />
                  <button id="erPdf" label="ZUGFeRD-PDF" imageMso="FileSaveAsPdfOrXps"
                          screentip="ZUGFeRD / Factur-X" supertip="Speichert ein PDF mit eingebetteter XML (ZUGFeRD 2.3, Profil EN 16931). Gilt als E-Rechnung im B2B-Bereich."
                          onAction="OnExportPdf" />
                </group>
                <group id="erReceive" label="Empfangen (Import)">
                  <button id="erImport" label="E-Rechnung öffnen" size="large" imageMso="FileOpen"
                          screentip="E-Rechnung öffnen" supertip="Liest XRechnung (UBL/CII) und ZUGFeRD/Factur-X-PDFs ein, prüft sie und trägt sie ins Blatt 'Rechnungseingang' ein. Mehrere Dateien gleichzeitig möglich."
                          onAction="OnImport" />
                  <button id="erInbox" label="Rechnungseingang" imageMso="ViewsDatasheetView"
                          screentip="Rechnungseingang" supertip="Zeigt die Liste aller eingelesenen E-Rechnungen."
                          onAction="OnInbox" />
                </group>
                <group id="erValidator" label="Offizielle Prüfung">
                  <checkBox id="erUseKosit" label="KoSIT-Prüfung"
                            screentip="KoSIT-Validator verwenden" supertip="Prüft zusätzlich mit dem offiziellen Validator der Koordinierungsstelle für IT-Standards (dauert ca. 10 Sekunden)."
                            getPressed="GetUseKosit" onAction="OnUseKosit" />
                  <button id="erKositStatus" label="Status / Installieren" imageMso="ControlProperties"
                          onAction="OnKositStatus" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;

    public void OnNew(IRibbonControl _) => Run(Commands.NewInvoice);
    public void OnDemo(IRibbonControl _) => Run(Commands.InsertDemo);
    public void OnCheck(IRibbonControl _) => Run(Commands.Validate);
    public void OnExportUbl(IRibbonControl _) => Run(() => Commands.Export(ExportKind.XRechnungUbl));
    public void OnExportCii(IRibbonControl _) => Run(() => Commands.Export(ExportKind.XRechnungCii));
    public void OnExportPdf(IRibbonControl _) => Run(() => Commands.Export(ExportKind.ZugferdPdf));
    public void OnImport(IRibbonControl _) => Run(Commands.Import);
    public void OnInbox(IRibbonControl _) => Run(Commands.ShowInbox);
    public void OnKositStatus(IRibbonControl _) => Run(Commands.KositStatus);

    public bool GetUseKosit(IRibbonControl _) => Settings.Current.UseKosit;
    public void OnUseKosit(IRibbonControl _, bool pressed) => Run(() => Commands.SetKositEnabled(pressed));

    private static void Run(Action action)
    {
        try
        {
            action();
        }
        catch (InvoiceDataException ex)
        {
            Ui.Error(ex.Message);
        }
        catch (Exception ex)
        {
            Ui.Error("Unerwarteter Fehler:\n\n" + ex.Message + "\n\n" + ex.GetType().Name + "\n" + ex.StackTrace);
        }
    }
}
