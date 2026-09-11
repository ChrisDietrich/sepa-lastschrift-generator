# CLAUDE.md

Hinweise für Claude Code beim Arbeiten in diesem Repo.

## Projekt

Python-Ersatz für ein VBA-Makro (`Excel2SepaXML.xlsm`, nicht Teil dieses
Repos), das ein Verein nutzte, um aus einer Mitgliederliste
SEPA-Lastschriften (pain.008.001.02-XML) für die Bank zu erzeugen, z. B.
für Kursgebühren-Einzüge. Grund für den Ersatz: abnehmende
Makro-Unterstützung in aktuellen Excel-Versionen.

## Architektur

- `sepa_lastschrift.py` - die gesamte Geschäftslogik: Validierung
  (IBAN-Prüfziffer, SEPA-Zeichensatz/-Längenregeln, Umlaut-Ersetzung),
  Einlesen der Mitgliederliste (Excel via `openpyxl`, oder CSV), Aufbau
  des SEPA-XML sowie eine CLI (`main()`/`parse_args()`). Zentraler
  Einstiegspunkt für beide Bedienwege ist `generate_sepa_file()` - CLI
  und GUI rufen ausschließlich diese Funktion auf, um Logik nicht zu
  duplizieren. Rückgabewert ist ein `GenerationSummary`
  (`results: list[GenerationResult]` + `skipped: list[tuple[str, str]]`),
  da eine Excel-Datei mit mehreren Tabellenblättern mehrere Ausgabedateien
  erzeugen kann (siehe unten).
- `gui.py` - Tkinter-Oberfläche, importiert `sepa_lastschrift` als Modul.
  Enthält keine eigene Geschäftslogik, nur Widgets und Fehleranzeige.
- Fehler werden über eigene Exceptions transportiert, nicht `SystemExit`,
  damit die GUI sie in Dialogen anzeigen kann:
  `ValidationError`, `SheetNotFoundError` (explizit per `--sheet`
  gewähltes Blatt existiert nicht), `RowValidationErrors` (trägt
  `.errors: list[str]`), `NoRowsError`, `NoValidSheetsError` (Automatik-
  Modus: kein einziges Tabellenblatt enthielt gültige Daten, trägt
  `.skipped: list[tuple[str, str]]`). CLI (`main()`) fängt sie ab und
  gibt sie auf stderr aus, GUI zeigt sie in `messagebox`-Dialogen.

## Mehrblatt-Verhalten (Automatik-Modus, kein `--sheet` angegeben)

Ohne explizit gewähltes Blatt wird jedes *sichtbare* Tabellenblatt einer
Excel-Datei versucht (`visible_sheet_names()` filtert versteckte Blätter
wie alte Nachschlagelisten aus dem Original-Template komplett heraus - die
werden nicht einmal als "übersprungen" gemeldet). Pro Blatt:

- keine Zeilen gefunden, oder Zeilen mit Validierungsfehlern → Blatt wird
  übersprungen, Grund landet in `GenerationSummary.skipped`, Verarbeitung
  der übrigen Blätter läuft weiter (Nutzer soll nicht durch ein einzelnes
  falsch beschriftetes Blatt blockiert werden)
- gültige Daten → eigene SEPA-Datei für dieses Blatt

Ist am Ende `results` leer, wird `NoValidSheetsError` geworfen (harter
Fehler - nichts wurde erzeugt). Wird dagegen explizit `--sheet` angegeben,
gilt das strikt: fehlt das Blatt → `SheetNotFoundError`, ist es leer/
fehlerhaft → `NoRowsError`/`RowValidationErrors` (kein Überspringen, da der
Nutzer die Auswahl bewusst getroffen hat). Gleiches gilt für CSV-Dateien
(kein Blattkonzept, immer strikt).

Message-/Payment-ID werden bei mehreren erzeugten Dateien pro Lauf mit
einem laufenden Index eindeutig gemacht (`<id>-<n>`), damit nicht mehrere
SEPA-Dateien mit identischer `MsgId` bei der Bank eingereicht werden.

## Wichtig: Faithfulness zum Original-Makro

Die Validierungsregeln (Regex für erlaubte Zeichen in IDs, Feldlängen wie
70/140/35 Zeichen, IBAN-Mod-97-Prüfziffer, Umlaut-zu-ASCII-Ersetzung vor
der Zeichenprüfung, "NOTPROVIDED" bei fehlender BIC/EndToEndId) sind
bewusst 1:1 aus dem VBA-Original übernommen, damit sich das Verhalten
gegenüber der bisherigen Excel-Lösung nicht ändert. Änderungen an diesen
Regeln nur nach expliziter Rücksprache, nicht "verbessern" ohne Anlass.

Erwartetes Spaltenformat der Mitgliederliste (Blattname beliebig, siehe
Mehrblatt-Verhalten unten; Zeile 1 = Kopfzeile, ab Zeile 2 Daten, Abbruch
bei leerer Spalte A):
`A Name | B Betrag | C BIC | D IBAN | E Verwendungszweck | F EndToEndId | G Mandatsreferenz | H Datum Mandatsunterschrift`

Standard ist Einzelbuchung (`batch_booking=False`), Sammelbuchung ist
Opt-in (CLI: `--batch-booking`, GUI: Checkbox unter „Erweiterte
Einstellungen"). Ausgabedateiname: `<Eingabedateiname ohne Endung>SEPA.xml`,
bei mehreren erzeugten Dateien pro Lauf zusätzlich mit `_<Blattname>` -
nicht mehr die alte `CDD_<MsgId>_<PmtInfId>.xml`-Namenskonvention.

## Vereinsdaten

`creditor.json` (Vereinsname/IBAN/Gläubiger-ID/BIC/Währung) enthält echte
Bankdaten und ist in `.gitignore` - niemals committen. `creditor.example.json`
ist die versionierte Vorlage mit Dummy-Daten. Im gebauten Programm liegt
`creditor.json` neben der `.exe`/`.app` (siehe `app_dir()` in `gui.py`,
wichtig wegen PyInstaller `--onefile`, das `__file__` sonst auf einen
flüchtigen Temp-Ordner zeigen lässt).

## Build & Release

`.github/workflows/build.yml` baut bei jedem Push auf `main` (und per
`workflow_dispatch`) mit PyInstaller je eine Windows-`.exe` und eine
macOS-`.app` und veröffentlicht beide als Assets im GitHub Release mit
festem Tag `latest` (kein Semantic Versioning, rollierendes Release - der
alte `latest`-Release/Tag wird vor jedem Lauf gelöscht und neu angelegt).
Download-Link für Endnutzer: `.../releases/latest`, nicht der
Actions-Artifacts-Bereich (erfordert GitHub-Login, daher ungeeignet für
Vereinsmitglieder ohne Account).

Lokal bauen: `pip install pyinstaller -r requirements.txt && pyinstaller
--onefile --windowed --name SEPA-Lastschrift-Generator gui.py`.

## Bekannte Stolperfalle: Tkinter auf macOS

Apples `/usr/bin/python3` bringt Tk 8.5 mit (uralt, rendert auf aktuellem
macOS oft nur ein leeres/graues Fenster). Für lokale GUI-Tests auf macOS
eine venv auf einem Python mit modernem Tk verwenden, z. B. Homebrew-Python
plus `brew install python-tk@<version>`. Der GitHub-Actions-Build ist
davon nicht betroffen (macos-latest-Runner hat ein funktionierendes Tk).

## Konventionen

- Sprache: durchgehend Deutsch (Code-Kommentare, Fehlermeldungen,
  CLI-Hilfetext, README, diese Datei) - Zielgruppe ist ein deutscher
  Verein.
- Keine Testdaten (`.xlsx`, `.xlsm`) oder generierten SEPA-XML-Dateien
  ins Repo - siehe `.gitignore`.
- Änderungen kurz in `CHANGELOG.md` unter `[Unreleased]` nachtragen.
- Es gibt (noch) keine automatisierten Tests - Änderungen an
  `sepa_lastschrift.py` manuell gegen eine Test-Mitgliederliste prüfen
  (IBAN-Validierung, Fehlerfälle, XML-Struktur).
