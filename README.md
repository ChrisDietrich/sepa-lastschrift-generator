# SEPA-Lastschrift-Generator in Python mit GUI

Erzeugt SEPA-Lastschriftdateien (pain.008.001.02-XML) aus einer Excel- oder
CSV-Mitgliederliste, z. B. für den Einzug von Kursgebühren oder
Mitgliedsbeiträgen. 

Zwei Bedienwege:

- **GUI** (`gui.py`) - Mitgliederliste auswählen, Fälligkeitstag eintragen,
  fertig. Für Kassenwart:innen ohne Python-Kenntnisse gedacht.
- **Kommandozeile** (`sepa_lastschrift.py`) - für Automatisierung/Skripte.


## Download (fertige Programme, ohne Python-Installation)

→ [Neueste Version](../../releases/latest) → passende Datei
herunterladen:

- **Windows:** `SEPA-Lastschrift-Generator-Windows.exe` direkt ausführen.
  Da die Datei nicht signiert ist, zeigt Windows beim ersten Start die
  SmartScreen-Warnung „Windows hat den Computer geschützt" - auf
  **„Weitere Informationen"** → **„Trotzdem ausführen"** klicken.
- **macOS:** `SEPA-Lastschrift-Generator-macOS.zip` entpacken, das
  enthaltene `.app` öffnen (ggf. Rechtsklick → Öffnen, da unsigniert).

Wird bei jedem Push automatisch von GitHub Actions neu gebaut und dort
aktualisiert - kein GitHub-Account zum Herunterladen nötig.

## Erste Schritte (GUI)

1. Programm starten (Windows: `.exe` doppelklicken, SmartScreen-Warnung
   bestätigen - siehe oben; macOS: `.app` öffnen, ggf. Rechtsklick →
   Öffnen, da unsigniert).
2. Beim ersten Start: **„Vereinsdaten..."** klicken und Vereinsname, IBAN
   und Gläubiger-ID eintragen. Werden danach dauerhaft neben dem Programm
   gespeichert (`creditor.json`).
3. **„Datei wählen..."** → Excel- oder CSV-Mitgliederliste auswählen.
4. Fälligkeitstag eintragen (Format `TT.MM.JJJJ`) oder **„+7 Tage"**
   klicken. Bitte anhand der Vorlagefrist eurer Bank prüfen, ob 7 Tage
   ausreichen.
5. **„SEPA-Datei erstellen"** klicken. Die XML-Datei landet im selben
   Ordner wie die Mitgliederliste.


## Format der Mitgliederliste

Spalten A-H (Kopfzeile in Zeile 1, Daten ab Zeile 2, Tabellenblattname
`SEPA_Lastschrift` bzw. per `--sheet` änderbar):

| Spalte | Inhalt                           | Pflicht |
|--------|----------------------------------|---------|
| A      | Name des Zahlungspflichtigen     | ja      |
| B      | Betrag                           | ja      |
| C      | BIC                              | nein    |
| D      | IBAN                             | ja      |
| E      | Verwendungszweck                 | ja      |
| F      | EndToEnd-Referenz                | nein    |
| G      | Mandatsreferenz                  | ja      |
| H      | Datum der Mandatsunterschrift    | ja      |


## Kommandozeile

```
python3 sepa_lastschrift.py mitglieder.xlsx --collection-date 01.10.2026
```

Alle Optionen samt ausführlicher Erläuterung (Lastschriftsequenz,
Lastschriftart, Fälligkeitstag/Vorlagefrist):

```
python3 sepa_lastschrift.py --help
```


## Entwicklung / lokale Installation

```
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python3 gui.py
```

Vereinsdaten für lokale Tests: `creditor.example.json` nach
`creditor.json` kopieren und ausfüllen. Diese Datei enthält echte
Kontodaten und ist bewusst in `.gitignore` ausgeschlossen.


## Windows-/macOS-Programm selbst bauen

```
pip install pyinstaller
pyinstaller --onefile --windowed --name SEPA-Lastschrift-Generator gui.py
```

Ergebnis liegt danach in `dist/`. Passiert automatisch für Windows und
macOS bei jedem Push auf `main` über
[`.github/workflows/build.yml`](.github/workflows/build.yml).

