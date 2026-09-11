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
  duplizieren.
- `gui.py` - Tkinter-Oberfläche, importiert `sepa_lastschrift` als Modul.
  Enthält keine eigene Geschäftslogik, nur Widgets und Fehleranzeige.
- Fehler werden über eigene Exceptions transportiert, nicht `SystemExit`,
  damit die GUI sie in Dialogen anzeigen kann:
  `ValidationError`, `SheetNotFoundError`, `RowValidationErrors` (trägt
  `.errors: list[str]`), `NoRowsError`. CLI (`main()`) fängt sie ab und
  gibt sie auf stderr aus, GUI zeigt sie in `messagebox`-Dialogen.

## Wichtig: Faithfulness zum Original-Makro

Die Validierungsregeln (Regex für erlaubte Zeichen in IDs, Feldlängen wie
70/140/35 Zeichen, IBAN-Mod-97-Prüfziffer, Umlaut-zu-ASCII-Ersetzung vor
der Zeichenprüfung, "NOTPROVIDED" bei fehlender BIC/EndToEndId) sind
bewusst 1:1 aus dem VBA-Original übernommen, damit sich das Verhalten
gegenüber der bisherigen Excel-Lösung nicht ändert. Änderungen an diesen
Regeln nur nach expliziter Rücksprache, nicht "verbessern" ohne Anlass.

Erwartetes Spaltenformat der Mitgliederliste (Blatt `SEPA_Lastschrift`,
Zeile 1 = Kopfzeile, ab Zeile 2 Daten, Abbruch bei leerer Spalte A):
`A Name | B Betrag | C BIC | D IBAN | E Verwendungszweck | F EndToEndId | G Mandatsreferenz | H Datum Mandatsunterschrift`

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
