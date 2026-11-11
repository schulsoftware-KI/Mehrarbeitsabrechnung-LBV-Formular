# Mehrarbeitsabrechnung-LBV-Formular

Excel-Makro mit Python-Hilfsprogrammen, das aus dem Untis-Export „Vertretungen / Entfälle“ die Mehrarbeit je Lehrkraft berechnet und die **LBV-Formulare (Änderungsmitteilung Mehrarbeitsvergütung, LBV(A/Bes)23_2019_02)** automatisch ausfüllt.

> **Hinweis:** Inoffizielles Hilfsmittel. Es ist keine Software des LBV oder eines Ministeriums. Alle Ergebnisse müssen vor dem Einreichen von Ihnen geprüft werden. Nutzung auf eigene Verantwortung, ohne Gewähr (siehe [LICENSE](LICENSE)).

## Was das Tool kann

- Untis-Export in die Excel-Datei importieren und je Lehrkraft Summen bilden (Entfälle, Vertretungen, Bagatellgrenze)
- Einzel-PDFs und Sammel-PDF der Abrechnung erstellen
- LBV-Formulare mit Barcode je Lehrkraft ausfüllen (`lbv_fill`)
- Mehrere Monatsformulare je Lehrkraft zu einem Gesamtformular vereinen (`lbv_merge`)
- Alle Einzel- oder alle Gesamtformulare zu **einer** PDF-Sammeldatei für alle Lehrkräfte zusammenfassen, mit Lesezeichen je Formular (`lbv_sammel`)
- Alle PDFs eines Ordners drucken, wahlweise über SumatraPDF, Adobe Acrobat oder das Standard-PDF-Programm, auf einen in der Arbeitsmappe festgelegten Drucker
- Abrechnungs-PDFs per SMTP (`lbv_email`) oder über Outlook versenden (Absenderkonto in der Arbeitsmappe festlegbar)

Die Namen der Lehrkräfte werden in der Arbeitsmappe über Pseudonyme (`L1`, `L2`, …) verarbeitet, damit die Untis-Daten nicht im Klartext herumliegen.

## Inhalt des Repositories

| Ordner | Inhalt |
|---|---|
| `beispiel/` | Excel-Beispieldatei mit ausschließlich erfundenen Daten |
| `python/` | Quellcode der vier Hilfsprogramme (`lbv_fill`, `lbv_merge`, `lbv_sammel`, `lbv_email`) |
| `vba/` | Exportierte VBA-Module, direkt in Excel importierbar (siehe unten) |
| `docs/` | Kurzanleitung |

> **Stand der Beispieldatei:** Die Makros in `beispiel/` können älter sein als die Module in `vba/`. Maßgeblich ist `vba/`. So aktualisieren Sie die Makros einer Arbeitsmappe: Alt+F11, die alten `Modul1` und `Modul2` entfernen (Rechtsklick → Entfernen, nicht exportieren), dann Datei → Datei importieren mit `vba/Modul1.bas` und `vba/Modul2.bas`. Danach in der Tabelle Informationen einmal „Buttons erstellen“ klicken.
>
> `vba/DieseArbeitsmappe1.bas` wird **nicht** importiert. Den Code daraus fügen Sie im VBA-Editor unter „Microsoft Excel Objekte“ in *DieseArbeitsmappe* ein, ab `Private Sub` bis `End Sub`. Nur dort funktioniert der Doppelklick in der Spalte „PDF erstellen“.

Die fertigen **EXE-Dateien** liegen nicht im Repository, sondern unter [Releases](../../releases).

## Voraussetzungen

- Windows und Microsoft Excel mit aktivierten Makros
- Entweder die EXE-Dateien aus dem Release **oder** Python 3 mit  
  `py -m pip install -r requirements.txt`  
  (bei der Python-Installation „Add Python to PATH“ ankreuzen)
- Die **leere LBV-Vorlage** (PDF) vom Landesamt für Besoldung und Versorgung. Sie ist hier bewusst nicht enthalten. Beziehen Sie die aktuelle Version selbst vom LBV und speichern Sie sie lokal.

## Einrichtung

1. Beispieldatei herunterladen. **Wichtig:** Rechtsklick auf die Datei → Eigenschaften → Haken bei „Zulassen“ → OK. Sonst blockiert Office die Makros.
2. `lbv_fill.exe`, `lbv_merge.exe`, `lbv_sammel.exe` und `lbv_email.exe` aus dem Release in einen gemeinsamen Ordner legen (oder die `.py`-Dateien, wenn Python installiert ist). Optional zum Drucken: die portable Version von [SumatraPDF](https://www.sumatrapdfreader.org) in denselben Ordner entpacken.
3. In der Tabelle **Informationen** eintragen: Jahreszahl, Dienststelle, Schulnummer, Bearbeiter, Telefon, Minuten je Unterrichtsstunde, Standardschlüssel, **Pfad und Dateiname der LBV-Vorlage**, **Pfad zu `lbv_fill.exe`** und je nach Bedarf:
   - `Outlook-Absender`: Adresse eines in Outlook eingerichteten Kontos (leer = Outlook-Standardkonto)
   - SMTP-Server, SMTP-Port, Verschlüsselung, Benutzername, Passwort für den Versand ohne Outlook
   - `Druckeradresse`: Name, `\\server\drucker` oder IP-Adresse eines installierten Druckers (leer = Standarddrucker)
   - `Druckprogramm`: `automatisch`, `SumatraPDF`, `Acrobat` oder `Standard-PDF-Programm` (leer = automatisch)
4. Die Tabelle **Pseudonyme** mit Kürzel, Nachname und Vorname Ihres Kollegiums füllen (die Beispieldatei enthält nur Platzhalter).

## Ablauf (Kurzfassung)

1. In Untis: Berichte → Vertretungsplanung → Vertretungsabrechnung, Datumsbereich des Abrechnungsmonats wählen, als Excel exportieren.
2. In der Arbeitsmappe **„Tabelle1 importieren“**, danach in der Übersicht **„Summen einlesen“**.
3. In der Spalte „PDF erstellen“ Lehrkräfte markieren (0/1 oder per Doppelklick).
4. **„PDFs erstellen“** bzw. **„LBV-Formular erstellen“** klicken.
5. Bei mehreren Monaten **„Merge LBV“**: erstellt je Lehrkraft `LBV_Mehrarbeit_<Kürzel>_Gesamt.pdf` und bietet danach direkt die Sammeldatei aller Gesamtformulare an.
6. **„LBV-Sammeldatei“**: fasst wahlweise alle Einzel- oder alle Gesamtformulare eines Ordners zu einer PDF zusammen. Die Formularfelder werden dabei fest in die Seiten übernommen, damit jede Seite die richtigen Werte zeigt.
7. **„PDFs drucken“**: druckt alle PDFs eines Ordners. Die E-Mail-Makros (`MehrarbeitPDFsVersenden` für Outlook, `MehrarbeitPDFsVersendenSMTP` für SMTP) starten Sie mit Alt+F8.

Ausführlich: [docs/Kurzanleitung Vertretungsabrechnung.docx](docs).

## Eigene EXE-Dateien bauen

```
py -m pip install pypdf reportlab pyinstaller
py -m PyInstaller --onefile --collect-submodules=reportlab.graphics.barcode lbv_fill.py
py -m PyInstaller --onefile --collect-submodules=reportlab.graphics.barcode lbv_merge.py
py -m PyInstaller --onefile lbv_sammel.py
py -m PyInstaller --onefile lbv_email.py
```

Ohne `--collect-submodules=reportlab.graphics.barcode` brechen `lbv_fill.exe` und `lbv_merge.exe` mit `No module named 'reportlab.graphics.barcode.code128'` ab.

Die fertige Datei liegt jeweils im Unterordner `dist`.

## Bekannte Einschränkungen und Sicherheitshinweise

- **Formularversion:** Die Barcode-Konstante (`ABES23/201902/`) gehört zur Formularversion 2019_02 und muss bei einer neuen Formularversion angepasst werden. Dasselbe gilt für die Feldnamen der Vorlage.
- **Teilzeit im Tarifbereich** (Bereich C des Formulars) muss in der Regel manuell geprüft bzw. bearbeitet werden.
- **Passwörter:** Das SMTP-Passwort steht im Klartext in der Tabelle *Informationen*. Beim Versand schreibt das Makro es außerdem vorübergehend in `%TEMP%\lbv_email.json`. Diese Datei wird nicht automatisch gelöscht. Nutzen Sie ein eigenes, nur dafür vorgesehenes Konto und löschen Sie die Datei nach dem Versand.
- **Datenschutz:** Die Arbeitsmappe verarbeitet personenbezogene Daten (Personalnummer, Geburtsdatum, Aktenzeichen). Speichern Sie die echte Arbeitsmappe nur an dafür freigegebenen Orten und **laden Sie niemals echte Daten in dieses Repository hoch**. Klären Sie den Einsatz mit Schulleitung und Datenschutzbeauftragten.
- **Virenscanner / SmartScreen:** Mit PyInstaller erzeugte EXE-Dateien sind nicht signiert und lösen häufig Fehlalarme aus. Sie können den Quellcode prüfen, die Prüfsummen im Release vergleichen oder die EXE-Dateien selbst bauen.

## Lizenz

[MIT](LICENSE)
